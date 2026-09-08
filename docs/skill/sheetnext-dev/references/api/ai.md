# AI API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [AI](#ai)

## AI
- Source: `src/core/AI/AI.js`
- AI assistant module (tool-calling workflow)

### Methods
#### `chatInput(con)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| con | string | Yes | - | - |

#### `async conversation(p, joinChat = true): Promise<void>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| p | string | Yes | - | - |
| joinChat | boolean | No | true | - |

**Returns**
- Type: `Promise<void>`

#### `clearChat()`

#### `handleFileChange(event)`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| event | Event | Yes | - | - |

#### `async screenshot(addresses = [])`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| addresses | Array | No | [] | - |
