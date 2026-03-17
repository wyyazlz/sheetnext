# Enum API

> Generated: 2026-03-17

## Enum Types

> Access via `SN.Enum.XXX`, or import from `sheetnext/enum`.

### EventPriority
- Event priority. Lower values run earlier.

| Key | Value | Description |
| --- | --- | --- |
| SYSTEM | -2000 | Internal system |
| PROTECTION | -1000 | Protection checks |
| PLUGIN | -500 | Plugin |
| NORMAL | 0 | Default |
| LATE | 500 | Post-processing |
| AUDIT | 1000 | Audit log |

### PasteMode
- Paste mode.

| Key | Value | Description |
| --- | --- | --- |
| ALL | all | All (default) |
| VALUE | value | Values only |
| FORMULA | formula | Formulas only |
| FORMAT | format | Formats only |
| NO_BORDER | noBorder | Without borders |
| COLUMN_WIDTH | colWidth | Column widths |
| COMMENT | comment | Comment |
| VALIDATION | validation | Data validation |
| TRANSPOSE | transpose | Transpose |
| VALUE_NUMBER_FORMAT | valueNumberFormat | Values and number formats |
| ALL_MERGE_CONDITIONAL | allMergeConditional | All with source theme |
| FORMULA_NUMBER_FORMAT | formulaNumberFormat | Formulas and number formats |

### PasteOperation
- Paste operation mode.

| Key | Value | Description |
| --- | --- | --- |
| NONE | none | None |
| ADD | add | Add |
| SUBTRACT | sub | Subtract |
| MULTIPLY | mul | Multiply |
| DIVIDE | div | Divide |

### AnchorType
- Drawing anchor type.

| Key | Value | Description |
| --- | --- | --- |
| TWO_CELL | twoCell | Move and resize with cells |
| ONE_CELL | oneCell | Move with cells only |
| ABSOLUTE | absolute | Fixed position |

### DrawingType
- Drawing type.

| Key | Value | Description |
| --- | --- | --- |
| CHART | chart | - |
| IMAGE | image | - |
| SHAPE | shape | - |
| SLICER | slicer | - |

### ValidationType
- Data validation type.

| Key | Value | Description |
| --- | --- | --- |
| LIST | list | Dropdown list |
| WHOLE | whole | Integer |
| DECIMAL | decimal | Decimal |
| DATE | date | Date |
| TIME | time | Time |
| TEXT_LENGTH | textLength | Text length |
| CUSTOM | custom | Custom formula |

### ValidationOperator
- Data validation operator.

| Key | Value | Description |
| --- | --- | --- |
| BETWEEN | between | Between |
| NOT_BETWEEN | notBetween | Not between |
| EQUAL | equal | Equal |
| NOT_EQUAL | notEqual | Not equal |
| GREATER_THAN | greaterThan | Greater than |
| GREATER_THAN_OR_EQUAL | greaterThanOrEqual | Greater than or equal |
| LESS_THAN | lessThan | Less than |
| LESS_THAN_OR_EQUAL | lessThanOrEqual | Less than or equal |

### FilterOperator
- Filter condition operator.

| Key | Value | Description |
| --- | --- | --- |
| EQUAL | equal | Equal |
| NOT_EQUAL | notEqual | Not equal |
| GREATER_THAN | greaterThan | Greater than |
| GREATER_THAN_OR_EQUAL | greaterThanOrEqual | Greater than or equal |
| LESS_THAN | lessThan | Less than |
| LESS_THAN_OR_EQUAL | lessThanOrEqual | Less than or equal |
| BEGINS_WITH | beginsWith | Begins with |
| ENDS_WITH | endsWith | Ends with |
| CONTAINS | contains | Contains |
| NOT_CONTAINS | notContains | Does not contain |

### CFType
- Conditional formatting type.

| Key | Value | Description |
| --- | --- | --- |
| CELL_VALUE | cellIs | Cell Value |
| TEXT_CONTAINS | containsText | Text contains |
| DATE_OCCURRING | timePeriod | Date occurring |
| TOP_BOTTOM | top10 | Top/bottom N |
| ABOVE_AVERAGE | aboveAverage | Above/below average |
| DUPLICATE | duplicateValues | Duplicate values |
| UNIQUE | uniqueValues | Unique values |
| DATA_BAR | dataBar | DATA BAR |
| COLOR_SCALE | colorScale | COLOUR STEP |
| ICON_SET | iconSet | ICON SET |
| EXPRESSION | expression | FORMULA |

### SortOrder
- Sort direction.

| Key | Value | Description |
| --- | --- | --- |
| ASC | asc | Ascending |
| DESC | desc | Descending |

### HorizontalAlign
- Horizontal alignment.

| Key | Value | Description |
| --- | --- | --- |
| LEFT | left | - |
| CENTER | center | - |
| RIGHT | right | - |
| JUSTIFY | justify | Justify |
| FILL | fill | Fill |
| DISTRIBUTED | distributed | Distributed |

### VerticalAlign
- Vertical alignment.

| Key | Value | Description |
| --- | --- | --- |
| TOP | top | - |
| CENTER | center | - |
| BOTTOM | bottom | - |
| JUSTIFY | justify | - |
| DISTRIBUTED | distributed | - |

### BorderStyle
- Border Style

| Key | Value | Description |
| --- | --- | --- |
| NONE | none | - |
| THIN | thin | - |
| MEDIUM | medium | - |
| THICK | thick | - |
| DASHED | dashed | - |
| DOTTED | dotted | - |
| DOUBLE | double | - |
