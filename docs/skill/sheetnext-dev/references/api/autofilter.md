# AutoFilter API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [AutoFilter](#autofilter)

## AutoFilter
- Source: `src/core/AutoFilter/AutoFilter.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | Sheet it belongs to |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| ref | string \| null | get/set | No | Gets/sets the default (sheet) filter range |
| sortState | {colIndex: number, order: 'asc' \| 'desc'} \| null | get/set | No | Default (sheet) scope sort state |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} \| null | No | Default (sheet) scope |
| columns | Map<number, Object> | No | Default (sheet) scope filter |
| enabled | boolean | No | Whether the default (sheet) scope is enabled |
| headerRow | number | No | Default (sheet) scope table header row |
| hasActiveFilters | boolean | No | Whether the default (sheet) scope has a valid filter |
| activeScopeId | string | No | Current Active Scope ID |

### Methods
#### `toggle(): void`

**Returns**
- Toggle the active table's filter buttons or the worksheet filter.

#### `setActiveScope(scopeId): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | Yes | - | - |

**Returns**
- Type: `string`
- Set active scope id when scope exists.

#### `getTableScopeId(tableId): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| tableId | string | Yes | - | - |

**Returns**
- Type: `string`
- Build table scope id from table id.

#### `isScopeEnabled(scopeId = null): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `boolean`
- Check whether scope range is enabled.

#### `hasActiveFiltersInScope(scopeId = null): boolean`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `boolean`
- Check whether scope has active column filters.

#### `hasAnyActiveFilters(): boolean`

**Returns**
- Type: `boolean`
- Check whether any scope has active filters.

#### `getScopeRange(scopeId = null): {s:{r:number,c:number},e:{r:number,c:number}} | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `{s:{r:number,c:number},e:{r:number,c:number}} | null`
- Get range of scope.

#### `getScopeHeaderRow(scopeId = null): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `number`
- Get header row index of scope.

#### `getColumnFilter(colIndex, scopeId = null): Object | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string \| null | No | null | - |

**Returns**
- Type: `Object | null`
- Get filter config for column in scope.

#### `getEnabledScopes(): Array<{scopeId:string,ownerType:string,ownerId:string | null,range:{s:{r:number,c:number},e:{r:number,c:number}},headerRow:number}>`

**Returns**
- Type: `Array<{scopeId:string,ownerType:string,ownerId:string | null,range:{s:{r:number,c:number},e:{r:number,c:number}},headerRow:number}>`
- List enabled scopes sorted by range.

#### `registerTableScope(table, options = {}): string | null`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| table | Object | Yes | - | - |
| options | Object | No | {} | - |

**Returns**
- Type: `string | null`
- Register or update filter scope for table.

#### `unregisterTableScope(tableId, options = {}): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| tableId | string | Yes | - | - |
| options | {silent?:boolean} | No | {} | - |

**Returns**
- Unregister table scope and refresh view state.

#### `setRange(range, scopeId = this._defaultScopeId)`
- Set filter scope (default sheet scope; scopeId can be specified)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| range | {s: {r: number, c: number}, e: {r: number, c: number}} | Yes | - | - |
| scopeId | string | No | this._defaultScopeId | - |

#### `isColumnInRange(colIndex, scopeId = null)`
- Determine if a column is within the filter range

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `isHeaderRow(rowIndex, scopeId = null)`
- Determine if a row is a table header row

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `getColumnValues(colIndex, scopeId = null): Array<{value: any, text: string, count: number, isEmpty: boolean}>`
- Get all unique values for a column (for filter panel)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

**Returns**
- Type: `Array<{value: any, text: string, count: number, isEmpty: boolean}>`

#### `setColumnFilter(colIndex, filter, scopeId = null)`
- Set column filters

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| filter | Object \| null | Yes | - | Null clears this column filter. |
| scopeId | string | No | null | - |
| filter.type | 'values' \| 'custom' | Yes | - | Filter mode. |
| filter.values | Array<string \| number \| boolean \| null> | No | - | Accepted values for values mode; an empty array clears the filter. |
| filter.filters | Array<{operator: string, value: string \| number}> | No | - | Conditions for custom mode, for example [{operator: 'contains', value: 'East'}]. |
| filter.and | boolean | No | true | Combine custom conditions with AND; false uses OR. |

**Examples**
```js
const filter = SN.activeSheet.AutoFilter;
filter.setRange('A1:D100');
filter.setColumnFilter(1, {
  type: 'custom',
  filters: [{ operator: 'contains', value: 'East' }],
  and: true
});
```

#### `clearColumnFilter(colIndex, scopeId = null)`
- Clear a column filter

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `clearAllFilters(scopeId = null)`
- Clear filters (reserved range)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | null | - |

#### `clear(scopeId = this._defaultScopeId)`
- Clear filters completely (inclusive)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | this._defaultScopeId | - |

#### `hasFilter(colIndex, scopeId = null)`
- Determine if columns are filtered

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `setSortState(colIndex, order, scopeId = null)`
- Set sorting status

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| order | 'asc' \| 'desc' | Yes | - | - |
| scopeId | string | No | null | - |

#### `getSortOrder(colIndex, scopeId = null)`
- Get sorting status

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| colIndex | number | Yes | - | - |
| scopeId | string | No | null | - |

#### `isRowFilteredHidden(rowIndex): boolean`
- Whether the row is hidden due to filtering

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rowIndex | number | Yes | - | - |

**Returns**
- Type: `boolean`

#### `parse(xmlObj)`
- Resolve default scopes from sheet AutoFilter

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | - |

#### `parseScope(scopeId, autoFilterXml, options = {})`
- Resolves the filter state of the specified scope to memory (mainly for table import)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | Yes | - | - |
| autoFilterXml | Object | Yes | - | - |
| options | {range?: Object, restoreHiddenRowsFromSheet?: boolean, ownerType?: string, ownerId?: string} | No | {} | - |

#### `buildXml(scopeId = this._defaultScopeId): Object | null`
- Build the XML for the specified scope; export the sheet scope by default

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| scopeId | string | No | this._defaultScopeId | - |

**Returns**
- Type: `Object | null`
