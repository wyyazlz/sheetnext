/**
 * 条件格式辅助函数
 * 包含颜色插值、条件比较、百分位计算等工具
 */

/**
 * 条件比较函数
 * @param {number} value - 待比较的值
 * @param {string} operator - 操作符
 * @param {number} formula1 - 比较值1
 * @param {number} formula2 - 比较值2（用于between）
 * @returns {boolean}
 */
export function compareValue(value, operator, formula1, formula2) {
    if (value === null || value === undefined) return false;

    switch (operator) {
        case 'greaterThan':
            return value > formula1;
        case 'lessThan':
            return value < formula1;
        case 'equal':
            return value === formula1;
        case 'notEqual':
            return value !== formula1;
        case 'greaterThanOrEqual':
            return value >= formula1;
        case 'lessThanOrEqual':
            return value <= formula1;
        case 'between':
            return value >= formula1 && value <= formula2;
        case 'notBetween':
            return value < formula1 || value > formula2;
        default:
            return false;
    }
}

/**
 * 解析颜色为RGB数组
 * @param {string} color - 颜色字符串 (#RRGGBB 或 rgb())
 * @returns {number[]} [r, g, b]
 */
export function parseColor(color) {
    if (!color) return [0, 0, 0];

    // #RRGGBB 格式
    if (color.startsWith('#')) {
        const hex = color.slice(1);
        if (hex.length === 6) {
            return [
                parseInt(hex.slice(0, 2), 16),
                parseInt(hex.slice(2, 4), 16),
                parseInt(hex.slice(4, 6), 16)
            ];
        }
        if (hex.length === 3) {
            return [
                parseInt(hex[0] + hex[0], 16),
                parseInt(hex[1] + hex[1], 16),
                parseInt(hex[2] + hex[2], 16)
            ];
        }
    }

    // rgb(r, g, b) 格式
    const match = color.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (match) {
        return [parseInt(match[1]), parseInt(match[2]), parseInt(match[3])];
    }

    return [0, 0, 0];
}

/**
 * RGB数组转十六进制颜色
 * @param {number[]} rgb - [r, g, b]
 * @returns {string} #RRGGBB
 */
export function rgbToHex(rgb) {
    return '#' + rgb.map(v =>
        Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')
    ).join('');
}

/**
 * 双色插值
 * @param {string} color1 - 起始颜色
 * @param {string} color2 - 结束颜色
 * @param {number} t - 插值比例 [0, 1]
 * @returns {string} 插值后的颜色
 */
export function interpolateColor(color1, color2, t) {
    const rgb1 = parseColor(color1);
    const rgb2 = parseColor(color2);

    const result = [
        rgb1[0] + (rgb2[0] - rgb1[0]) * t,
        rgb1[1] + (rgb2[1] - rgb1[1]) * t,
        rgb1[2] + (rgb2[2] - rgb1[2]) * t
    ];

    return rgbToHex(result);
}

/**
 * 三色插值
 * @param {string} color1 - 起始颜色
 * @param {string} color2 - 中间颜色
 * @param {string} color3 - 结束颜色
 * @param {number} t - 插值比例 [0, 1]
 * @returns {string} 插值后的颜色
 */
export function interpolateColor3(color1, color2, color3, t) {
    if (t <= 0.5) {
        return interpolateColor(color1, color2, t * 2);
    } else {
        return interpolateColor(color2, color3, (t - 0.5) * 2);
    }
}

/**
 * 使颜色变亮
 * @param {string} color - 原始颜色
 * @param {number} amount - 变亮程度 [0, 1]
 * @returns {string}
 */
export function lightenColor(color, amount) {
    const rgb = parseColor(color);
    const result = rgb.map(v => v + (255 - v) * amount);
    return rgbToHex(result);
}

/**
 * 使颜色变暗
 * @param {string} color - 原始颜色
 * @param {number} amount - 变暗程度 [0, 1]
 * @returns {string}
 */
export function darkenColor(color, amount) {
    const rgb = parseColor(color);
    const result = rgb.map(v => v * (1 - amount));
    return rgbToHex(result);
}

/**
 * 计算百分位值
 * @param {number[]} sortedValues - 已排序的数值数组
 * @param {number} percentile - 百分位 [0, 100]
 * @returns {number}
 */
export function getPercentile(sortedValues, percentile) {
    if (sortedValues.length === 0) return 0;
    if (sortedValues.length === 1) return sortedValues[0];

    const index = (percentile / 100) * (sortedValues.length - 1);
    const lower = Math.floor(index);
    const upper = Math.ceil(index);
    const weight = index - lower;

    if (lower === upper) return sortedValues[lower];
    return sortedValues[lower] * (1 - weight) + sortedValues[upper] * weight;
}

/**
 * 计算标准差
 * @param {number[]} values - 数值数组
 * @param {number} avg - 平均值
 * @returns {number}
 */
export function getStdDev(values, avg) {
    if (values.length === 0) return 0;
    const squaredDiffs = values.map(v => Math.pow(v - avg, 2));
    const avgSquaredDiff = squaredDiffs.reduce((a, b) => a + b, 0) / values.length;
    return Math.sqrt(avgSquaredDiff);
}

/**
 * 根据cfvo配置计算阈值
 * @param {Object} cfvo - cfvo配置 { type, val }
 * @param {Object} rangeData - 范围数据
 * @param {Function} resolveFormula - 公式解析函数
 * @returns {number}
 */
export function resolveCfvoValue(cfvo, rangeData, resolveFormula) {
    if (!cfvo) return 0;

    const { sorted, min, max, avg, count } = rangeData;
    const numValues = sorted.map(s => s.val);

    switch (cfvo.type) {
        case 'min':
            return min;
        case 'max':
            return max;
        case 'num':
            return cfvo.val ?? 0;
        case 'percent':
            // 百分比：min + (max - min) * percent / 100
            return min + (max - min) * (cfvo.val ?? 0) / 100;
        case 'percentile':
            return getPercentile(numValues, cfvo.val ?? 50);
        case 'formula':
            return resolveFormula ? resolveFormula(cfvo.val) : 0;
        case 'autoMin':
            // 自动最小值（用于数据条负值）
            return Math.min(0, min);
        case 'autoMax':
            // 自动最大值（用于数据条负值）
            return Math.max(0, max);
        default:
            return 0;
    }
}

/**
 * 判断是否为错误值
 * @param {*} value - 单元格值
 * @returns {boolean}
 */
export function isErrorValue(value) {
    if (typeof value !== 'string') return false;
    const errors = ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#GETTING_DATA', '#SPILL!', '#CALC!'];
    return errors.includes(value);
}

/**
 * 判断是否为空白值
 * @param {*} value - 单元格值
 * @returns {boolean}
 */
export function isBlankValue(value) {
    return value === null || value === undefined || value === '';
}

/**
 * 获取今天的起始时间戳
 * @returns {number}
 */
export function getTodayStart() {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
}

/**
 * 获取本周的起始时间戳（周日为第一天）
 * @returns {number}
 */
export function getWeekStart() {
    const now = new Date();
    const day = now.getDay();
    return getTodayStart() - day * 86400000;
}

/**
 * 获取本月的起始时间戳
 * @returns {number}
 */
export function getMonthStart() {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1).getTime();
}

/**
 * Excel日期序列号转时间戳
 * @param {number} serial - Excel日期序列号
 * @returns {number} 时间戳
 */
export function excelDateToTimestamp(serial) {
    // Excel日期基准：1900-01-01
    // 但Excel错误地认为1900年是闰年，所以要减1
    const excelEpoch = new Date(1899, 11, 30).getTime();
    return excelEpoch + serial * 86400000;
}

/**
 * 时间戳转Excel日期序列号
 * @param {number} timestamp - 时间戳
 * @returns {number} Excel日期序列号
 */
export function timestampToExcelDate(timestamp) {
    const excelEpoch = new Date(1899, 11, 30).getTime();
    return (timestamp - excelEpoch) / 86400000;
}

/**
 * 扩展十六进制颜色为8位（AARRGGBB）
 * @param {string} hex - 6位颜色 #RRGGBB
 * @returns {string} 8位颜色 FFRRGGBB
 */
export function expandHex(hex) {
    if (!hex) return 'FF000000';
    const h = hex.startsWith('#') ? hex.slice(1) : hex;
    return 'FF' + h.toUpperCase();
}
