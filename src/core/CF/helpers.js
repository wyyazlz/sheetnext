/**
 * Conditional formatting helper function
 * Contains tools for color interpolation, conditional comparisons, percentile calculations, and more
 */

/**
 * Conditional formatting helper function
 * Contains tools for color interpolation, conditional comparisons, percentile calculations, and more
 */

/**
 * Conditional Comparison Function
 * @param {number} value - Value to compare
 * @param {string} operator - Operator
 * @param {number} formula1 - Comparison value 1
 * @param {number} formula2 - Comparison value 2 (for between)
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
 * Resolve Color to RGB Array
 * @param {string} color - Color string (# RRGGBB or rgb ())
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
 * RGB array to hex color
 * @param {number[]} rgb - [r, g, b]
 * @returns {string} #RRGGBB
 */
export function rgbToHex(rgb) {
    return '#' + rgb.map(v =>
        Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')
    ).join('');
}

/**
 * Two-color interpolation
 * @param {string} color1 - Start color
 * @param {string} color2 - End Color
 * @param {number} t - Interpolation Scale [0, 1]
 * @ returns {string} interpolated color
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
 * Tricolor interpolation
 * @param {string} color1 - Start color
 * @param {string} color2 - Middle Color
 * @param {string} color3 - End Color
 * @param {number} t - Interpolation Scale [0, 1]
 * @ returns {string} interpolated color
 */
export function interpolateColor3(color1, color2, color3, t) {
    if (t <= 0.5) {
        return interpolateColor(color1, color2, t * 2);
    } else {
        return interpolateColor(color2, color3, (t - 0.5) * 2);
    }
}

/**
 * Brighten Colors
 * @param {string} color - Original Color
 * @param {number} amount - Brightness [0, 1]
 * @returns {string}
 */
export function lightenColor(color, amount) {
    const rgb = parseColor(color);
    const result = rgb.map(v => v + (255 - v) * amount);
    return rgbToHex(result);
}

/**
 * Darken Colors
 * @param {string} color - Original Color
 * @param {number} amount - Degree of darkening [0, 1]
 * @returns {string}
 */
export function darkenColor(color, amount) {
    const rgb = parseColor(color);
    const result = rgb.map(v => v * (1 - amount));
    return rgbToHex(result);
}

/**
 * Calculate Percentile
 * @param {number[]} sortedValues - Sorted Numeric Array
 * @param {number} percentile - Percentile [0, 100]
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
 * Calculate standard deviation
 * @param {number[]} values - Numeric Array
 * @param {number} avg - Mean
 * @returns {number}
 */
export function getStdDev(values, avg) {
    if (values.length === 0) return 0;
    const squaredDiffs = values.map(v => Math.pow(v - avg, 2));
    const avgSquaredDiff = squaredDiffs.reduce((a, b) => a + b, 0) / values.length;
    return Math.sqrt(avgSquaredDiff);
}

/**
 * Calculate thresholds based on cfvo configuration
 * @param {Object} cfvo - cfvo configuration {type, val}
 * @param {Object} rangeData - Scope data
 * @param {Function} resolveFormula - Formula parsing function
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
 * Determine if it is an error value
 * @param {*} value - Cell Value
 * @returns {boolean}
 */
export function isErrorValue(value) {
    if (typeof value !== 'string') return false;
    const errors = ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#GETTING_DATA', '#SPILL!', '#CALC!'];
    return errors.includes(value);
}

/**
 * Determine if it is a blank value
 * @param {*} value - Cell Value
 * @returns {boolean}
 */
export function isBlankValue(value) {
    return value === null || value === undefined || value === '';
}

/**
 * Get today's start timestamp
 * @returns {number}
 */
export function getTodayStart() {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
}

/**
 * Get the start timestamp of the week (Sunday is the first day)
 * @returns {number}
 */
export function getWeekStart() {
    const now = new Date();
    const day = now.getDay();
    return getTodayStart() - day * 86400000;
}

/**
 * Get the start timestamp of the month
 * @returns {number}
 */
export function getMonthStart() {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1).getTime();
}

/**
 * Excel date sequence number to timestamp
 * @param {number} serial - Excel date serial number
 * @ returns {number} timestamps
 */
export function excelDateToTimestamp(serial) {
    // Excel日期基准：1900-01-01
    // 但Excel错误地认为1900年是闰年，所以要减1
    const excelEpoch = new Date(1899, 11, 30).getTime();
    return excelEpoch + serial * 86400000;
}

/**
 * Timestamp to Excel date serial number
 * @param {number} timestamp - Timestamp
 * @ returns {number} Excel date serial number
 */
export function timestampToExcelDate(timestamp) {
    const excelEpoch = new Date(1899, 11, 30).getTime();
    return (timestamp - excelEpoch) / 86400000;
}

/**
 * Extended hex color is 8 bits (AARRGGBB)
 * @param {string} hex - 6 bit color # RRGGBB
 * @ returns {string} 8-bit color FFRRGGBB
 */
export function expandHex(hex) {
    if (!hex) return 'FF000000';
    const h = hex.startsWith('#') ? hex.slice(1) : hex;
    return 'FF' + h.toUpperCase();
}
