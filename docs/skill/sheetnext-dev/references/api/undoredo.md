# UndoRedo API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [UndoRedo](#undoredo)

## UndoRedo
- Source: `src/core/UndoRedo/UndoRedo.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| maxStack | number | No | 50 | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| canUndo | boolean | No | - |
| canRedo | boolean | No | - |

### Methods
#### `add(action): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| action | {undo: Function, redo: Function, activeCell?: {r:number,c:number}} | Yes | - | - |

**Returns**
- Add undo action into current transaction.

#### `begin(name = ''): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | No | '' | - |

**Returns**
- Begin a manual undo transaction.

#### `commit(): void`

**Returns**
- Commit current manual transaction.

#### `undo(): boolean`

**Returns**
- Type: `boolean`

#### `redo(): boolean`

**Returns**
- Type: `boolean`

#### `clear(): void`

**Returns**
- Clear undo/redo stacks and pending transaction.
