# Table API

> Generated: 2026-09-08
> SheetNext source package: 0.2.23; source fingerprint: 896bd17c14ed91f3295183303ef16d4d78bb335764d4362f17ab30915bafee95

[API index](../core-api.md)

- [Table](#table)
- [TableItem](#tableitem)

## Table
- Source: `src/core/Table/Table.js`

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

#### `appendRow(tableOrId): boolean`
- Append one blank data row to a table without shifting worksheet rows.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| tableOrId | TableItem \| string | Yes | - | - |

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

## TableItem
- Source: `src/core/Table/TableItem.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| id | string | No | options.id \|\| this._generateId() | - |
| columns | Array<Object> | No | options.columns \|\| [] | - |
| showHeaderRow | boolean | No | options.showHeaderRow ?? true | - |
| showFirstColumn | boolean | No | options.showFirstColumn ?? false | - |
| showLastColumn | boolean | No | options.showLastColumn ?? false | - |
| showRowStripes | boolean | No | options.showRowStripes ?? true | - |
| showColumnStripes | boolean | No | options.showColumnStripes ?? false | - |
| headerRowDxfId | number \| undefined | No | options.headerRowDxfId | - |
| dataDxfId | number \| undefined | No | options.dataDxfId | - |
| headerRowBorderDxfId | number \| undefined | No | options.headerRowBorderDxfId | - |
| tableBorderDxfId | number \| undefined | No | options.tableBorderDxfId | - |
| totalsRowBorderDxfId | number \| undefined | No | options.totalsRowBorderDxfId | - |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| name | string | get/set | No | - |
| displayName | string | get/set | No | - |
| styleId | string \| null | get/set | No | - |
| showTotalsRow | boolean | get/set | No | - |
| autoFilterEnabled | boolean | get/set | No | - |
| ref | string \| null | get/set | No | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} \| null | No | - |
| headerRowIndex | number | No | - |
| dataStartRow | number | No | - |
| dataEndRow | number | No | - |
| totalsRowIndex | number | No | - |
| style | Object \| null | No | - |

### Methods
#### `contains(row, col): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `isHeaderCell(row, col): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `isTotalsCell(row, col): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `isDataCell(row, col): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |
| col | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `getDataRowIndex(row): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| row | number | Yes | - | - |

**Returns**
- Type: `number`
- visible row index in data area, -1 if hidden or outside

#### `getColumnIndex(col): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| col | number | Yes | - | - |

**Returns**
- Type: `number`

#### `getColumn(colIndex): Object | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |

**Returns**
- Type: `Object | null`

#### `getColumnIndexByName(name): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `number`

#### `resize(newRef)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| newRef | string \| Object | Yes | - | - |

#### `syncColumnNames()`
- Sync column names from header cell values.

#### `applyTotalsRow(options = {})`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | {initializeDefaults?: boolean, force?: boolean} | No | {} | - |

#### `setTotalsFunction(colIndex, funcType)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| funcType | string \| null | Yes | - | - |

#### `convertToRange()`
- Convert table to plain range.

#### `toJSON(): Object`

**Returns**
- Type: `Object`

#### `static fromJSON(sheet, json): TableItem`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | Yes | - | - |
| json | Object | Yes | - | - |

**Returns**
- Type: `TableItem`
