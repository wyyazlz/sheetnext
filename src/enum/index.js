/**
 * SheetNext 常量枚举
 * 通过 SN.Enum 访问
 */

/**
 * 事件优先级（越小越先执行）
 */
export const EventPriority = Object.freeze({
    SYSTEM: -2000,      // 系统内部
    PROTECTION: -1000,  // 保护检查
    PLUGIN: -500,       // 插件
    NORMAL: 0,          // 默认
    LATE: 500,          // 后置处理
    AUDIT: 1000         // 审计日志
});

/**
 * 粘贴模式
 */
export const PasteMode = Object.freeze({
    ALL: 'all',                           // 全部（默认）
    VALUE: 'value',                       // 仅值
    FORMULA: 'formula',                   // 仅公式
    FORMAT: 'format',                     // 仅格式
    NO_BORDER: 'noBorder',                // 无边框
    COLUMN_WIDTH: 'colWidth',             // 列宽
    COMMENT: 'comment',                   // 批注
    VALIDATION: 'validation',             // 数据验证
    TRANSPOSE: 'transpose',               // 转置
    VALUE_NUMBER_FORMAT: 'valueNumberFormat',       // 值和数字格式
    ALL_MERGE_CONDITIONAL: 'allMergeConditional',   // 全部使用源主题
    FORMULA_NUMBER_FORMAT: 'formulaNumberFormat'    // 公式和数字格式
});

/**
 * 运算粘贴模式
 */
export const PasteOperation = Object.freeze({
    NONE: 'none',       // 无
    ADD: 'add',         // 加
    SUBTRACT: 'sub',    // 减
    MULTIPLY: 'mul',    // 乘
    DIVIDE: 'div'       // 除
});

/**
 * 图纸锚点类型
 */
export const AnchorType = Object.freeze({
    TWO_CELL: 'twoCell',    // 随单元格移动和缩放
    ONE_CELL: 'oneCell',    // 仅随单元格移动
    ABSOLUTE: 'absolute'    // 固定位置
});

/**
 * 图纸类型
 */
export const DrawingType = Object.freeze({
    CHART: 'chart',
    IMAGE: 'image',
    SHAPE: 'shape',
    SLICER: 'slicer'
});

/**
 * 数据验证类型
 */
export const ValidationType = Object.freeze({
    LIST: 'list',               // 下拉列表
    WHOLE: 'whole',             // 整数
    DECIMAL: 'decimal',         // 小数
    DATE: 'date',               // 日期
    TIME: 'time',               // 时间
    TEXT_LENGTH: 'textLength',  // 文本长度
    CUSTOM: 'custom'            // 自定义公式
});

/**
 * 数据验证操作符
 */
export const ValidationOperator = Object.freeze({
    BETWEEN: 'between',                     // 介于
    NOT_BETWEEN: 'notBetween',              // 未介于
    EQUAL: 'equal',                         // 等于
    NOT_EQUAL: 'notEqual',                  // 不等于
    GREATER_THAN: 'greaterThan',            // 大于
    GREATER_THAN_OR_EQUAL: 'greaterThanOrEqual',  // 大于等于
    LESS_THAN: 'lessThan',                  // 小于
    LESS_THAN_OR_EQUAL: 'lessThanOrEqual'   // 小于等于
});

/**
 * 筛选条件操作符
 */
export const FilterOperator = Object.freeze({
    EQUAL: 'equal',                         // 等于
    NOT_EQUAL: 'notEqual',                  // 不等于
    GREATER_THAN: 'greaterThan',            // 大于
    GREATER_THAN_OR_EQUAL: 'greaterThanOrEqual',  // 大于等于
    LESS_THAN: 'lessThan',                  // 小于
    LESS_THAN_OR_EQUAL: 'lessThanOrEqual',  // 小于等于
    BEGINS_WITH: 'beginsWith',              // 开头是
    ENDS_WITH: 'endsWith',                  // 结尾是
    CONTAINS: 'contains',                   // 包含
    NOT_CONTAINS: 'notContains'             // 不包含
});

/**
 * 条件格式类型
 */
export const CFType = Object.freeze({
    CELL_VALUE: 'cellIs',           // 单元格值
    TEXT_CONTAINS: 'containsText',  // 包含文本
    DATE_OCCURRING: 'timePeriod',   // 发生日期
    TOP_BOTTOM: 'top10',            // 前/后 N 项
    ABOVE_AVERAGE: 'aboveAverage',  // 高于/低于平均值
    DUPLICATE: 'duplicateValues',   // 重复值
    UNIQUE: 'uniqueValues',         // 唯一值
    DATA_BAR: 'dataBar',            // 数据条
    COLOR_SCALE: 'colorScale',      // 色阶
    ICON_SET: 'iconSet',            // 图标集
    EXPRESSION: 'expression'        // 公式
});

/**
 * 排序方向
 */
export const SortOrder = Object.freeze({
    ASC: 'asc',     // 升序
    DESC: 'desc'    // 降序
});

/**
 * 对齐方式 - 水平
 */
export const HorizontalAlign = Object.freeze({
    LEFT: 'left',
    CENTER: 'center',
    RIGHT: 'right',
    JUSTIFY: 'justify',     // 两端对齐
    FILL: 'fill',           // 填充
    DISTRIBUTED: 'distributed'  // 分散对齐
});

/**
 * 对齐方式 - 垂直
 */
export const VerticalAlign = Object.freeze({
    TOP: 'top',
    CENTER: 'center',
    BOTTOM: 'bottom',
    JUSTIFY: 'justify',
    DISTRIBUTED: 'distributed'
});

/**
 * 边框样式
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

// 统一导出
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
