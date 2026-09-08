# Cell API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Cell](#cell)

## Cell
- Source: `src/core/Cell/Cell.js`
- Cell classes

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| cIndex | number | No | cIndex | Column Index |
| row | Row | No | row | Rows belonging to |
| isMerged | boolean | No | false | Whether or not to merge cells |
| master | {r:number, c:number} \| null | No | null | Merge Cell Main Cell References |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| editVal | string | get/set | No | Edit value or formula |
| richText | Array \| null | get/set | No | Rich text runs array |
| hyperlink | Object \| null | get/set | No | Hyperlink Configuration |
| dataValidation | Object \| null | get/set | No | Data Validation Configuration |
| protection | {locked: boolean, hidden: boolean} | get/set | No | Cell Protection Configuration |
| control | Object \| null | get/set | No | Cell control configuration. |
| type | string | get/set | No | Cell Type |
| style | Object | get/set | No | Cell Styles |
| border | Object | get/set | No | Border Style |
| fill | Object | get/set | No | Fill Style |
| font | Object | get/set | No | Font style |
| alignment | Object | get/set | No | Alignment |
| numFmt | string \| undefined | get/set | No | Number Format |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| showVal | string | No | Show values |
| accountingData | Object \| undefined | No | Get Accounting Format Structured Data (for Rendering) |
| calcVal | any | No | Calculated value |
| spillArray | Array[] \| null | No | Get full spill array results |
| isSpillSource | boolean | No | Whether it is a spill source cell |
| isSpillRef | boolean | No | Whether it is a spill reference cell |
| horizontalAlign | string | No | Calculated Horizontal Alignment |
| verticalAlign | string | No | Calculated Vertical Alignment |
| isLocked | boolean | No | Is it locked |
| validData | boolean | No | Data validation results |
| isFormula | boolean | No | Is it a formula |

### Methods
#### `validateData(value = this.calcVal): boolean`
- Check a candidate calculated value against this cell's data validation rule.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| value | any | No | this.calcVal | Candidate calculated value |

**Returns**
- Type: `boolean`
