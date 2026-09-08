# Formula API

> Generated: 2026-09-08
> SheetNext source package: 0.2.24; source fingerprint: dc22deb262eb5bb1514a446034bb678f6f32c276259b681cc707c7f3c2e52752

[API index](../core-api.md)

- [Formula](#formula)
- [FormulaAstParser](#formulaastparser)
- [FormulaError](#formulaerror)
- [FormulaLexer](#formulalexer)
- [FormulaParseError](#formulaparseerror)

## Formula
- Source: `src/core/Formula/Formula.js`
- Formula Calculator

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| SN | - | No | SN | - |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| lastCalcVolatile | boolean | No | Whether the last calculation was volatile |
| lastResultIsDate | boolean | No | Whether the last calculated result is a date type |

### Methods
#### `calcFormula(formulaStr, cell): any`
- Calculation Formula

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| formulaStr | string | Yes | - | Equation String (without Equals) |
| cell | Cell | Yes | - | Current Cell |

**Returns**
- Type: `any`

#### `parseDeps(formulaOrAst, defaultSheetName = null): {sheetName:string | null,r:number,c:number}[]`
- Extract all dependent cell coordinates from a formula AST/string.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| formulaOrAst | string \| Object | Yes | - | Formula string without equals or parsed AST |
| defaultSheetName | string | No | null | Sheet name used by local references |

**Returns**
- Type: `{sheetName:string | null,r:number,c:number}[]`

#### `parseDependencyAreas(formulaOrAst, defaultSheetName = null): {sheetName:string | null,s:{r:number,c:number},e:{r:number,c:number}}[]`
- Extract dependency areas without expanding large ranges.

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| formulaOrAst | string \| Object | Yes | - | Formula string without equals or parsed AST |
| defaultSheetName | string | No | null | Sheet name used by local references |

**Returns**
- Type: `{sheetName:string | null,s:{r:number,c:number},e:{r:number,c:number}}[]`

#### `executeFunction(funcName, stack): any`
- Execute Function

**Parameters**
| Name | Type | Required | Default | Description |
| --- | --- | --- | --- | --- |
| funcName | string | Yes | - | Function name |
| stack | Array | Yes | - | Parameter Stack |

**Returns**
- Type: `any`

### Potential Lifecycle Props
- Assigned outside the constructor. These are flagged for audit because they may be public state that the scanner should not silently miss.
| Name | Static | Source Line | Example Assignment |
| --- | --- | --- | --- |
| currentCell | No | 81 | cell |

## FormulaAstParser
- Source: `src/core/Formula/FormulaAst.js`

### Methods
#### `parse()`

## FormulaError
- Source: `src/core/Formula/FormulaValue.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | No | 'FormulaError' | - |
| code | - | No | this.message | - |

## FormulaLexer
- Source: `src/core/Formula/FormulaAst.js`

### Methods
#### `scan()`

## FormulaParseError
- Source: `src/core/Formula/FormulaAst.js`

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| name | string | No | 'FormulaParseError' | - |
