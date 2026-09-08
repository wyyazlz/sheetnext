# IO API

> Generated: 2026-09-08
> SheetNext source package: 0.2.23; source fingerprint: 896bd17c14ed91f3295183303ef16d4d78bb335764d4362f17ab30915bafee95

[API index](../core-api.md)

- [IO](#io)

## IO
- Source: `src/core/IO/IO.js`

### Methods
#### `async import(file): Promise<void>`
- Import an xlsx/csv/json file. Some failures are caught and shown as notifications;<br>promise fulfillment is not a reliable success signal. Confirm the imported result before persisting.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| file | File | Yes | - | - |

**Returns**
- Type: `Promise<void>`

#### `async importFromUrl(fileUrl): Promise<void>`
- Fetch and import a file URL. Requires browser CORS access; failures may be handled by editor notifications.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| fileUrl | string | Yes | - | - |

**Returns**
- Type: `Promise<void>`
- Completion only; no reliable success/failure result.

#### `async export(type)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| type | 'XLSX' \| 'CSV' \| 'JSON' \| 'HTML' | Yes | - | - |

#### `async exportAllImage(): Promise<void>`

**Returns**
- Type: `Promise<void>`

#### `async getData(): Promise<Object>`
- Source: `src/core/IO/JsonIO.js` | Category: `JsonIO`

**Returns**
- Type: `Promise<Object>`

#### `setData(data): boolean`
- Replace the workbook with JSON version 2.0 data and clear undo/redo history on success.<br>Loading is not atomic: false may leave partially replaced state. Back up with getData before loading,<br>pause persistence during replacement, and reapply application permissions after loading.
- Source: `src/core/IO/JsonIO.js` | Category: `JsonIO`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| data | Object | Yes | - | Workbook JSON previously exported with getData; requires a non-empty sheets array. |

**Returns**
- Type: `boolean`
- True on completion; false on invalid input or a restore failure. No automatic rollback.
