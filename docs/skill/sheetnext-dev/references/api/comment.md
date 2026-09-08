# Comment API

> Generated: 2026-09-08
> SheetNext source package: 0.2.23; source fingerprint: 896bd17c14ed91f3295183303ef16d4d78bb335764d4362f17ab30915bafee95

[API index](../core-api.md)

- [Comment](#comment)
- [CommentItem](#commentitem)

## Comment
- Source: `src/core/Comment/Comment.js`
- Comment Manager<br>High-performance design: Map index (cellRef - > CommentItem) + O (1) lookup<br>Architectural Reference: Fully Replicated Drawing Design Patterns

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| map | Map<string, CommentItem> | No | new Map() | - |

### Methods
#### `parse(xmlObj)`
- Parsing comment data from xmlObj (imported from an Excel file)

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| xmlObj | Object | Yes | - | XML Object for Sheet |

#### `add(config = {}): CommentItem`
- Add Comment

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| config | Object | No | {} | {cellRef, author, text, visible, ...} |

**Returns**
- Type: `CommentItem`

#### `get(cellRef): CommentItem | null`
- Get comments

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string | Yes | - | Cell reference "A1" |

**Returns**
- Type: `CommentItem | null`

#### `remove(cellRef): boolean`
- Delete Comment

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellRef | string | Yes | - | Cell reference "A1" |

**Returns**
- Type: `boolean`

#### `getAll(): CommentItem[]`
- Get a list of all comments

**Returns**
- Type: `CommentItem[]`

#### `clearCache()`
- Clear the location cache for all comment (called when the view changes)

#### `toXmlObject(): Object | null`
- Export as XML object (for Excel export)<br>Convert thread comment to normal comment format

**Returns**
- Type: `Object | null`

## CommentItem
- Source: `src/core/Comment/CommentItem.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | - |
| id | string | No | config.id ?? `comment_${Date.now()}_${Math.random()}` | - |
| cellRef | string | No | config.cellRef | - |
| author | string | No | - | - |
| text | string | No | config.text ?? '' | - |
| visible | boolean | No | config.visible ?? false | - |
| width | number | No | config.width ?? 144 | - |
| height | number | No | config.height ?? 72 | - |
| marginLeft | number | No | config.marginLeft ?? 107 | - |
| marginTop | number | No | config.marginTop ?? 7 | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| cellPos | CellNum | No | - |
