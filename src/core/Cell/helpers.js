// ==================== Cell 辅助函数 ====================

// 数据验证两数比较
export const compareValue = (val, operator, formula1, formula2) => {
    const v1 = Number(formula1);
    const v2 = Number(formula2);
    switch (operator) {
        case 'between': return v2 !== null && val >= v1 && val <= v2;
        case 'notBetween': return v2 !== null && (val < v1 || val > v2);
        case 'equal': return val === v1;
        case 'notEqual': return val !== v1;
        case 'greaterThan': return val > v1;
        case 'lessThan': return val < v1;
        case 'greaterThanOrEqual': return val >= v1;
        case 'lessThanOrEqual': return val <= v1;
        default: return true;
    }
}

// 颜色标准格式化
export function expandHex(color) {
    if (typeof color !== 'string') return color;
    const s = color.charAt(0) === '#' ? color.slice(1) : color;
    if (s.length === 3) return 'FF' + s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
    if (s.length === 6) return 'FF' + s;
    return color;
}

// 从 NumFmt 模块重新导出，保持兼容性
export { numFmtFun, dateStrToDate, dateTrans, dateToSerial, modifyDecimalPlaces, formatValue } from './NumFmt.js';
