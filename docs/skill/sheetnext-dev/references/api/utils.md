# Utils API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Utils](#utils)

## Utils
- Source: `src/core/Utils/Utils.js`

### Methods
#### `objToArr(obj): Array<any>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | any | Yes | - | - |

**Returns**
- Type: `Array<any>`
- Normalize any value to array.

#### `rgbToHex(rgb): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| rgb | string | Yes | - | - |

**Returns**
- Type: `string`
- Convert rgb() string to hex color.

#### `emuToPx(emu): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| emu | number | Yes | - | - |

**Returns**
- Type: `number`
- Convert EMU to pixels.

#### `pxToEmu(px): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| px | number | Yes | - | - |

**Returns**
- Type: `number`
- Convert pixels to EMU.

#### `decodeEntities(encodedString): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| encodedString | string | Yes | - | - |

**Returns**
- Type: `string`
- Decode HTML entities to plain text.

#### `sortObjectInPlace(obj, keyOrder): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | Object | Yes | - | - |
| keyOrder | Array<string> | Yes | - | - |

**Returns**
- Reorder object keys in-place.

#### `debounce(func, wait): Function`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| func | Function | Yes | - | - |
| wait | number | Yes | - | - |

**Returns**
- Type: `Function`
- Create a debounced function.

#### `numToChar(num): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| num | number | Yes | - | - |

**Returns**
- Type: `string`

#### `charToNum(name): number`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | Yes | - | - |

**Returns**
- Type: `number`

#### `rangeNumToStr(obj, absolute = false): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| obj | {s:{r:number,c:number},e:{r:number,c:number}} | Yes | - | - |
| absolute | boolean | No | false | - |

**Returns**
- Type: `string`

#### `cellStrToNum(cellStr): {c:number,r:number} | undefined`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellStr | string | Yes | - | - |

**Returns**
- Type: `{c:number,r:number} | undefined`

#### `cellNumToStr(cellNum): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| cellNum | {c:number,r:number} | Yes | - | - |

**Returns**
- Type: `string`

#### `getFillFormula(formula, originalRow, originalCol, targetRow, targetCol): string`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| formula | string | Yes | - | - |
| originalRow | number | Yes | - | - |
| originalCol | number | Yes | - | - |
| targetRow | number | Yes | - | - |
| targetCol | number | Yes | - | - |

**Returns**
- Type: `string`
- Build fill formula by shifting relative refs only.

#### `toast(message): void`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| message | any | Yes | - | - |

**Returns**
- Show toast message.

#### `async downloadImageToBase64(imageUrl): Promise<string | undefined>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| imageUrl | string | Yes | - | - |

**Returns**
- Type: `Promise<string | undefined>`
- Fetch image URL and return base64 data URL.

#### `modal(options = {}): Promise<{bodyEl: HTMLElement} | null>`

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| options | Object | No | {} | - |

**Returns**
- Type: `Promise<{bodyEl: HTMLElement} | null>`
- Open common modal panel.
