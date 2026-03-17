/**
 * SheetNext enum constants.
 * Access via `SN.Enum`.
 */

/**
 * Event priority. Lower values run earlier.
 */
export const EventPriority = Object.freeze({
    SYSTEM: -2000,      // Internal system
    PROTECTION: -1000,  // Protection checks
    PLUGIN: -500,       // Plugin
    NORMAL: 0,          // Default
    LATE: 500,          // Post-processing
    AUDIT: 1000         // Audit log
});

/**
 * Paste mode.
 */
export const PasteMode = Object.freeze({
    ALL: 'all',                           // All (default)
    VALUE: 'value',                       // Values only
    FORMULA: 'formula',                   // Formulas only
    FORMAT: 'format',                     // Formats only
    NO_BORDER: 'noBorder',                // Without borders
    COLUMN_WIDTH: 'colWidth',             // Column widths
    COMMENT: 'comment',                   // Comment
    VALIDATION: 'validation',             // Data validation
    TRANSPOSE: 'transpose',               // Transpose
    VALUE_NUMBER_FORMAT: 'valueNumberFormat',       // Values and number formats
    ALL_MERGE_CONDITIONAL: 'allMergeConditional',   // All with source theme
    FORMULA_NUMBER_FORMAT: 'formulaNumberFormat'    // Formulas and number formats
});

/**
 * Paste operation mode.
 */
export const PasteOperation = Object.freeze({
    NONE: 'none',       // None
    ADD: 'add',         // Add
    SUBTRACT: 'sub',    // Subtract
    MULTIPLY: 'mul',    // Multiply
    DIVIDE: 'div'       // Divide
});

/**
 * Drawing anchor type.
 */
export const AnchorType = Object.freeze({
    TWO_CELL: 'twoCell',    // Move and resize with cells
    ONE_CELL: 'oneCell',    // Move with cells only
    ABSOLUTE: 'absolute'    // Fixed position
});

/**
 * Drawing type.
 */
export const DrawingType = Object.freeze({
    CHART: 'chart',
    IMAGE: 'image',
    SHAPE: 'shape',
    SLICER: 'slicer'
});

/**
 * Data validation type.
 */
export const ValidationType = Object.freeze({
    LIST: 'list',               // Dropdown list
    WHOLE: 'whole',             // Integer
    DECIMAL: 'decimal',         // Decimal
    DATE: 'date',               // Date
    TIME: 'time',               // Time
    TEXT_LENGTH: 'textLength',  // Text length
    CUSTOM: 'custom'            // Custom formula
});

/**
 * Data validation operator.
 */
export const ValidationOperator = Object.freeze({
    BETWEEN: 'between',                     // Between
    NOT_BETWEEN: 'notBetween',              // Not between
    EQUAL: 'equal',                         // Equal
    NOT_EQUAL: 'notEqual',                  // Not equal
    GREATER_THAN: 'greaterThan',            // Greater than
    GREATER_THAN_OR_EQUAL: 'greaterThanOrEqual',  // Greater than or equal
    LESS_THAN: 'lessThan',                  // Less than
    LESS_THAN_OR_EQUAL: 'lessThanOrEqual'   // Less than or equal
});

/**
 * Filter condition operator.
 */
export const FilterOperator = Object.freeze({
    EQUAL: 'equal',                         // Equal
    NOT_EQUAL: 'notEqual',                  // Not equal
    GREATER_THAN: 'greaterThan',            // Greater than
    GREATER_THAN_OR_EQUAL: 'greaterThanOrEqual',  // Greater than or equal
    LESS_THAN: 'lessThan',                  // Less than
    LESS_THAN_OR_EQUAL: 'lessThanOrEqual',  // Less than or equal
    BEGINS_WITH: 'beginsWith',              // Begins with
    ENDS_WITH: 'endsWith',                  // Ends with
    CONTAINS: 'contains',                   // Contains
    NOT_CONTAINS: 'notContains'             // Does not contain
});

/**
 * Conditional formatting type.
 */
export const CFType = Object.freeze({
    CELL_VALUE: 'cellIs',           // Cell Value
    TEXT_CONTAINS: 'containsText',  // Text contains
    DATE_OCCURRING: 'timePeriod',   // Date occurring
    TOP_BOTTOM: 'top10',            // Top/bottom N
    ABOVE_AVERAGE: 'aboveAverage',  // Above/below average
    DUPLICATE: 'duplicateValues',   // Duplicate values
    UNIQUE: 'uniqueValues',         // Unique values
    DATA_BAR: 'dataBar',            // DATA BAR
    COLOR_SCALE: 'colorScale',      // COLOUR STEP
    ICON_SET: 'iconSet',            // ICON SET
    EXPRESSION: 'expression'        // FORMULA
});

/**
 * Sort direction.
 */
export const SortOrder = Object.freeze({
    ASC: 'asc',     // Ascending
    DESC: 'desc'    // Descending
});

/**
 * Horizontal alignment.
 */
export const HorizontalAlign = Object.freeze({
    LEFT: 'left',
    CENTER: 'center',
    RIGHT: 'right',
    JUSTIFY: 'justify',     // Justify
    FILL: 'fill',           // Fill
    DISTRIBUTED: 'distributed'  // Distributed
});

/**
 * Vertical alignment.
 */
export const VerticalAlign = Object.freeze({
    TOP: 'top',
    CENTER: 'center',
    BOTTOM: 'bottom',
    JUSTIFY: 'justify',
    DISTRIBUTED: 'distributed'
});

/**
 * Border Style
 */
export const BorderStyle = Object.freeze({
    NONE: 'none',
    THIN: 'thin',
    MEDIUM: 'medium',
    THICK: 'thick',
    DASHED: 'dashed',
    DOTTED: 'dotted',
    DOUBLE: 'double'
});

// Unified export
export default {
    EventPriority,
    PasteMode,
    PasteOperation,
    AnchorType,
    DrawingType,
    ValidationType,
    ValidationOperator,
    FilterOperator,
    CFType,
    SortOrder,
    HorizontalAlign,
    VerticalAlign,
    BorderStyle
};
