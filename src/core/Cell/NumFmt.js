/**
 * Excel 数值格式化模块
 * 完整实现 Excel numFmt 规范，支持：
 * - 4段格式（正数;负数;零;文本）
 * - 条件格式 [>100]、颜色 [Red]、locale [$-804]
 * - 日期时间（自动区分 m 月/分）
 * - 数字（千分位、小数、百分比、科学计数）
 * - 分数（自动约分、固定分母）
 * - 特殊格式（填充、对齐、文本占位）
 */

import { BUILTIN_NUMFMTS } from './constants.js';

// ==================== 常量 ====================

const COLORS = {
    'black': '#000000', 'red': '#FF0000', 'green': '#00FF00', 'blue': '#0000FF',
    'yellow': '#FFFF00', 'magenta': '#FF00FF', 'cyan': '#00FFFF', 'white': '#FFFFFF',
    'color1': '#000000', 'color2': '#FFFFFF', 'color3': '#FF0000', 'color4': '#00FF00',
    'color5': '#0000FF', 'color6': '#FFFF00', 'color7': '#FF00FF', 'color8': '#00FFFF',
};

const CONDITION_OPS = { '=': (a, b) => a === b, '<>': (a, b) => a !== b, '>': (a, b) => a > b, '<': (a, b) => a < b, '>=': (a, b) => a >= b, '<=': (a, b) => a <= b };

// ==================== 缓存 ====================

const formatCache = new Map();
const CACHE_MAX = 200;

function getCachedAST(fmt) {
    if (formatCache.has(fmt)) return formatCache.get(fmt);
    const ast = parseFormat(fmt);
    if (formatCache.size >= CACHE_MAX) {
        const firstKey = formatCache.keys().next().value;
        formatCache.delete(firstKey);
    }
    formatCache.set(fmt, ast);
    return ast;
}

// ==================== 词法分析 ====================

const TokenType = {
    LITERAL: 1, CONDITION: 2, COLOR: 3, LOCALE: 4,
    DIGIT_0: 10, DIGIT_HASH: 11, DIGIT_Q: 12,
    DECIMAL: 13, THOUSAND: 14, PERCENT: 15, EXPONENT: 16, FRACTION: 17,
    DATE_Y: 20, DATE_M: 21, DATE_D: 22, TIME_H: 23, TIME_S: 24, TIME_ELAPSED: 25, AMPM: 26,
    FILL: 30, SKIP: 31, AT: 32,
};

function tokenize(fmt) {
    const tokens = [];
    let i = 0, len = fmt.length;

    while (i < len) {
        const ch = fmt[i];

        // 方括号内容 [...]
        if (ch === '[') {
            const end = fmt.indexOf(']', i);
            if (end === -1) { tokens.push({ type: TokenType.LITERAL, value: ch }); i++; continue; }
            const content = fmt.slice(i + 1, end);
            const lower = content.toLowerCase();

            // 条件 [>100] [=5]
            const condMatch = content.match(/^([<>=]+)(-?\d+\.?\d*)$/);
            if (condMatch) {
                tokens.push({ type: TokenType.CONDITION, op: condMatch[1], value: parseFloat(condMatch[2]) });
            }
            // 颜色 [Red] [Color1]
            else if (COLORS[lower]) {
                tokens.push({ type: TokenType.COLOR, value: COLORS[lower] });
            }
            // locale [$-804] [$¥-804]
            else if (content.startsWith('$')) {
                const localeMatch = content.match(/^\$([^-]*)-?(.*)$/);
                tokens.push({ type: TokenType.LOCALE, symbol: localeMatch?.[1] || '', code: localeMatch?.[2] || '' });
            }
            // 累计时间 [h] [m] [s]
            else if (/^[hms]{1,2}$/i.test(content)) {
                tokens.push({ type: TokenType.TIME_ELAPSED, unit: lower[0], len: content.length });
            }
            // DBNum 等忽略
            else {
                tokens.push({ type: TokenType.LITERAL, value: '' });
            }
            i = end + 1;
            continue;
        }

        // 引号字符串 "..."
        if (ch === '"') {
            const end = fmt.indexOf('"', i + 1);
            if (end === -1) { i++; continue; }
            tokens.push({ type: TokenType.LITERAL, value: fmt.slice(i + 1, end) });
            i = end + 1;
            continue;
        }

        // 转义字符 \x
        if (ch === '\\' && i + 1 < len) {
            tokens.push({ type: TokenType.LITERAL, value: fmt[i + 1] });
            i += 2;
            continue;
        }

        // 填充 *x
        if (ch === '*' && i + 1 < len) {
            tokens.push({ type: TokenType.FILL, value: fmt[i + 1] });
            i += 2;
            continue;
        }

        // 跳过宽度 _x
        if (ch === '_' && i + 1 < len) {
            tokens.push({ type: TokenType.SKIP, value: fmt[i + 1] });
            i += 2;
            continue;
        }

        // AM/PM
        const ampmMatch = fmt.slice(i).match(/^(AM\/PM|A\/P)/i);
        if (ampmMatch) {
            tokens.push({ type: TokenType.AMPM, value: ampmMatch[0] });
            i += ampmMatch[0].length;
            continue;
        }

        // 日期时间 token
        const dtMatch = fmt.slice(i).match(/^(yyyy|yy|y|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s)/i);
        if (dtMatch) {
            const v = dtMatch[0].toLowerCase();
            let type;
            if (v[0] === 'y') type = TokenType.DATE_Y;
            else if (v[0] === 'm') type = TokenType.DATE_M; // 后续 Parser 区分月/分
            else if (v[0] === 'd') type = TokenType.DATE_D;
            else if (v[0] === 'h') type = TokenType.TIME_H;
            else if (v[0] === 's') type = TokenType.TIME_S;
            tokens.push({ type, value: v, len: v.length });
            i += dtMatch[0].length;
            continue;
        }

        // 科学计数 E+ E-
        if ((ch === 'E' || ch === 'e') && i + 1 < len && (fmt[i + 1] === '+' || fmt[i + 1] === '-')) {
            tokens.push({ type: TokenType.EXPONENT, sign: fmt[i + 1] });
            i += 2;
            continue;
        }

        // 单字符 token
        switch (ch) {
            case '0': tokens.push({ type: TokenType.DIGIT_0 }); break;
            case '#': tokens.push({ type: TokenType.DIGIT_HASH }); break;
            case '?': tokens.push({ type: TokenType.DIGIT_Q }); break;
            case '.': tokens.push({ type: TokenType.DECIMAL }); break;
            case ',': tokens.push({ type: TokenType.THOUSAND }); break;
            case '%': tokens.push({ type: TokenType.PERCENT }); break;
            case '/': tokens.push({ type: TokenType.FRACTION }); break;
            case '@': tokens.push({ type: TokenType.AT }); break;
            case ' ': tokens.push({ type: TokenType.LITERAL, value: ' ' }); break;
            case ':': tokens.push({ type: TokenType.LITERAL, value: ':' }); break;
            case '-': tokens.push({ type: TokenType.LITERAL, value: '-' }); break;
            case '(': case ')': tokens.push({ type: TokenType.LITERAL, value: ch }); break;
            case '年': case '月': case '日': case '时': case '分': case '秒':
                tokens.push({ type: TokenType.LITERAL, value: ch }); break;
            default:
                // 其他字符作为字面量
                if (/[^\s]/.test(ch)) tokens.push({ type: TokenType.LITERAL, value: ch });
                break;
        }
        i++;
    }

    return tokens;
}

// ==================== 语法分析 ====================

function parseFormat(fmt) {
    if (!fmt || fmt === 'General') return { sections: [{ type: 'general', tokens: [] }] };

    // 分割段（注意引号和方括号内的分号不分割）
    const sections = [];
    let current = '', inQuote = false, inBracket = false;
    for (let i = 0; i < fmt.length; i++) {
        const ch = fmt[i];
        if (ch === '"') inQuote = !inQuote;
        else if (ch === '[') inBracket = true;
        else if (ch === ']') inBracket = false;
        if (ch === ';' && !inQuote && !inBracket) {
            sections.push(current);
            current = '';
        } else {
            current += ch;
        }
    }
    sections.push(current);

    return {
        sections: sections.map((sec, idx) => parseSection(sec, idx, sections.length))
    };
}

function parseSection(fmt, index, total) {
    const tokens = tokenize(fmt);
    const section = { tokens: [], condition: null, color: null, locale: null, type: 'number' };

    // 提取元信息
    const contentTokens = [];
    for (const tok of tokens) {
        if (tok.type === TokenType.CONDITION) section.condition = { op: tok.op, value: tok.value };
        else if (tok.type === TokenType.COLOR) section.color = tok.value;
        else if (tok.type === TokenType.LOCALE) {
            section.locale = tok;
            if (tok.symbol) contentTokens.push({ type: TokenType.LITERAL, value: tok.symbol });
        }
        else contentTokens.push(tok);
    }

    // 检测格式类型
    const hasDate = contentTokens.some(t => [TokenType.DATE_Y, TokenType.DATE_D].includes(t.type));
    const hasTime = contentTokens.some(t => [TokenType.TIME_H, TokenType.TIME_S, TokenType.TIME_ELAPSED, TokenType.AMPM].includes(t.type));
    const hasFraction = contentTokens.some(t => t.type === TokenType.FRACTION);
    const hasAt = contentTokens.some(t => t.type === TokenType.AT);

    if (hasAt && !hasDate && !hasTime) section.type = 'text';
    else if (hasDate || hasTime) section.type = 'datetime';
    else if (hasFraction) section.type = 'fraction';
    else section.type = 'number';

    // 解析 m/mm 歧义（在日期时间格式中）
    if (section.type === 'datetime') {
        resolveDateTimeM(contentTokens);
    }

    section.tokens = contentTokens;

    // 默认条件分配
    if (!section.condition && total <= 4) {
        if (total === 1) section.condition = null; // 应用所有
        else if (total === 2) section.condition = index === 0 ? { op: '>=', value: 0 } : { op: '<', value: 0 };
        else if (total >= 3) {
            if (index === 0) section.condition = { op: '>', value: 0 };
            else if (index === 1) section.condition = { op: '<', value: 0 };
            else if (index === 2) section.condition = { op: '=', value: 0 };
        }
    }

    return section;
}

// 解析 m/mm 是月份还是分钟
function resolveDateTimeM(tokens) {
    for (let i = 0; i < tokens.length; i++) {
        if (tokens[i].type !== TokenType.DATE_M) continue;

        let isMinute = false;

        // 向前查找
        for (let j = i - 1; j >= 0; j--) {
            const t = tokens[j].type;
            if (t === TokenType.TIME_H || t === TokenType.TIME_ELAPSED) { isMinute = true; break; }
            if (t === TokenType.DATE_Y || t === TokenType.DATE_D) break;
            if ([TokenType.DIGIT_0, TokenType.DIGIT_HASH].includes(t)) break;
        }

        // 向后查找
        if (!isMinute) {
            for (let j = i + 1; j < tokens.length; j++) {
                const t = tokens[j].type;
                if (t === TokenType.TIME_S) { isMinute = true; break; }
                if (t === TokenType.DATE_Y || t === TokenType.DATE_D) break;
                if ([TokenType.DIGIT_0, TokenType.DIGIT_HASH].includes(t)) break;
            }
        }

        if (isMinute) tokens[i].isMinute = true;
    }
}

// ==================== 格式化执行 ====================

export function formatValue(value, fmt, cellType) {
    // 空值处理
    if (value == null || value === '') return '';

    // General 格式
    if (!fmt || fmt === 'General') {
        if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', color: null };
        if (typeof value === 'number') return { text: formatGeneral(value), color: null };
        return { text: String(value), color: null };
    }

    // 内置格式转换
    if (typeof fmt === 'number') fmt = BUILTIN_NUMFMTS[fmt] || 'General';

    const ast = getCachedAST(fmt);
    const numValue = typeof value === 'number' ? value : parseFloat(value);
    const isNum = !isNaN(numValue);

    // 非数值：有文本段则用文本段格式化，否则原样返回
    if (!isNum) {
        const textSection = ast.sections.find(s => s.type === 'text');
        if (textSection) return { text: formatText(String(value), textSection.tokens), color: textSection.color };
        return { text: String(value), color: null };
    }

    // 选择适用的段
    let section = ast.sections[0];
    if (ast.sections.length > 1) {
        for (const sec of ast.sections) {
            if (sec.type === 'text') continue;
            if (!sec.condition || CONDITION_OPS[sec.condition.op]?.(numValue, sec.condition.value)) {
                section = sec;
                break;
            }
        }
    }

    // 根据类型格式化
    let result;
    switch (section.type) {
        case 'general':
            result = formatGeneral(numValue);
            break;
        case 'text':
            result = formatText(String(value), section.tokens);
            break;
        case 'datetime':
            result = formatDateTime(numValue, section.tokens, cellType);
            break;
        case 'fraction':
            result = formatFraction(numValue, section.tokens);
            break;
        default:
            result = formatNumber(numValue, section.tokens);
    }

    // 会计格式返回结构化数据
    if (result && result.isAccounting) {
        return {
            text: result.prefix + result.number + result.suffix, // 兼容 showVal
            color: section.color,
            accounting: result // 结构化数据供渲染使用
        };
    }

    return { text: result, color: section.color };
}

// 简化版导出（只返回文本）
export function numFmtFun(value, fmt, cellType) {
    const res = formatValue(value, fmt, cellType);
    return typeof res === 'object' ? res.text : res;
}

// ==================== General 格式 ====================

function formatGeneral(value) {
    if (typeof value !== 'number') return String(value ?? '');
    if (!isFinite(value)) return value > 0 ? '∞' : '-∞';
    if (value === 0) return '0';

    const abs = Math.abs(value);
    // 科学计数法阈值
    if (abs >= 1e11 || (abs < 1e-9 && abs !== 0)) {
        return value.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E');
    }
    // 去除尾随零
    let str = value.toPrecision(10);
    if (str.includes('.')) str = str.replace(/\.?0+$/, '');
    return str;
}

// ==================== 文本格式 ====================

function formatText(value, tokens) {
    let result = '';
    for (const tok of tokens) {
        if (tok.type === TokenType.AT) result += value;
        else if (tok.type === TokenType.LITERAL) result += tok.value;
        else if (tok.type === TokenType.SKIP) result += ' ';
    }
    return result || value;
}

// ==================== 日期时间格式 ====================

function formatDateTime(serial, tokens, cellType) {
    if (isNaN(serial)) return '';

    const date = serialToDate(serial);
    if (!date) return String(serial);

    // Excel 基准日期用于累计时间
    const baseDate = new Date(1899, 11, 30, 0, 0, 0, 0);
    const totalMs = serial * 24 * 60 * 60 * 1000;

    let result = '';
    let hasAMPM = tokens.some(t => t.type === TokenType.AMPM);
    let hour = date.getHours();
    let hour12 = hour % 12 || 12;
    let ampm = hour < 12 ? 'AM' : 'PM';

    for (const tok of tokens) {
        switch (tok.type) {
            case TokenType.DATE_Y:
                if (tok.len >= 4) result += date.getFullYear();
                else if (tok.len >= 2) result += String(date.getFullYear() % 100).padStart(2, '0');
                else result += date.getFullYear() % 100;
                break;

            case TokenType.DATE_M:
                if (tok.isMinute) {
                    // 分钟
                    result += tok.len >= 2 ? String(date.getMinutes()).padStart(2, '0') : date.getMinutes();
                } else {
                    // 月份
                    const month = date.getMonth() + 1;
                    if (tok.len >= 4) result += ['', '一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'][month];
                    else if (tok.len === 3) result += ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][month];
                    else if (tok.len === 2) result += String(month).padStart(2, '0');
                    else result += month;
                }
                break;

            case TokenType.DATE_D:
                const day = date.getDate();
                const dow = date.getDay();
                if (tok.len >= 4) result += ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'][dow];
                else if (tok.len === 3) result += ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][dow];
                else if (tok.len === 2) result += String(day).padStart(2, '0');
                else result += day;
                break;

            case TokenType.TIME_H:
                const h = hasAMPM ? hour12 : hour;
                result += tok.len >= 2 ? String(h).padStart(2, '0') : h;
                break;

            case TokenType.TIME_S:
                result += tok.len >= 2 ? String(date.getSeconds()).padStart(2, '0') : date.getSeconds();
                break;

            case TokenType.TIME_ELAPSED:
                const totalH = Math.floor(totalMs / 3600000);
                const totalM = Math.floor(totalMs / 60000);
                const totalS = Math.floor(totalMs / 1000);
                if (tok.unit === 'h') result += tok.len >= 2 ? String(totalH).padStart(2, '0') : totalH;
                else if (tok.unit === 'm') result += tok.len >= 2 ? String(totalM).padStart(2, '0') : totalM;
                else if (tok.unit === 's') result += tok.len >= 2 ? String(totalS).padStart(2, '0') : totalS;
                break;

            case TokenType.AMPM:
                result += tok.value.includes('/') ? (hour < 12 ? tok.value.split('/')[0] : tok.value.split('/')[1]) : ampm;
                break;

            case TokenType.LITERAL:
                result += tok.value;
                break;

            case TokenType.SKIP:
                result += ' ';
                break;

            case TokenType.DECIMAL:
                result += '.';
                break;

            case TokenType.DIGIT_0:
                // 毫秒部分（.0 或 .00 或 .000）
                const ms = date.getMilliseconds();
                result += String(Math.floor(ms / 100));
                break;

            case TokenType.FRACTION:
                // 日期格式中的 / 分隔符
                result += '/';
                break;
        }
    }

    return result;
}

// Excel 序列号转 JS Date
function serialToDate(serial) {
    if (typeof serial !== 'number' || isNaN(serial)) return null;

    // 处理 1900 年 bug（Excel 错误认为 1900 年是闰年）
    let days = serial;
    if (serial > 60) days -= 1; // 跳过虚假的 1900/2/29
    else if (serial === 60) days = 59; // 1900/2/28

    const baseDate = new Date(1899, 11, 30);
    const wholeDays = Math.floor(days);
    const timeFraction = days - wholeDays;

    const result = new Date(baseDate);
    result.setDate(result.getDate() + wholeDays);

    if (timeFraction > 0) {
        const ms = Math.round(timeFraction * 24 * 60 * 60 * 1000);
        result.setTime(result.getTime() + ms);
    }

    return result;
}

// JS Date 转 Excel 序列号
export function dateToSerial(date) {
    if (!(date instanceof Date)) return date;

    const baseDate = new Date(1899, 11, 30);
    const diffMs = date.getTime() - baseDate.getTime();
    let days = diffMs / (24 * 60 * 60 * 1000);

    // 1900 年 bug 修正
    if (days >= 60) days += 1;

    return days;
}

// ==================== 数字格式 ====================

function formatNumber(value, tokens) {
    if (isNaN(value)) return String(value);

    let absValue = Math.abs(value);
    let isNegative = value < 0;
    let result = '';

    // 分析格式
    let hasPercent = tokens.some(t => t.type === TokenType.PERCENT);
    let hasExponent = tokens.some(t => t.type === TokenType.EXPONENT);
    let hasFill = tokens.some(t => t.type === TokenType.FILL);
    let scaleCount = 0; // 逗号缩放（千、百万等）

    // 统计尾部逗号缩放
    let foundDecimalOrDigit = false;
    for (let i = tokens.length - 1; i >= 0; i--) {
        const t = tokens[i];
        if ([TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q, TokenType.DECIMAL].includes(t.type)) {
            foundDecimalOrDigit = true;
            break;
        }
        if (t.type === TokenType.THOUSAND && !foundDecimalOrDigit) scaleCount++;
    }

    // 应用百分比
    if (hasPercent) absValue *= 100;

    // 应用缩放
    if (scaleCount > 0) absValue /= Math.pow(1000, scaleCount);

    // 科学计数法
    if (hasExponent) {
        result = formatExponent(absValue, tokens);
    } else {
        result = formatDecimal(absValue, tokens, hasFill);
    }

    // 处理负数括号
    const hasParens = tokens.some(t => t.type === TokenType.LITERAL && t.value === '(');
    if (isNegative && !hasParens) {
        // 检查是否有负号字面量
        const hasNegLiteral = tokens.some(t => t.type === TokenType.LITERAL && t.value === '-');
        if (!hasNegLiteral) {
            // 对于会计格式结构化数据，负号加到 number 部分
            if (result && result.isAccounting) {
                result.number = '-' + result.number;
            } else {
                result = '-' + result;
            }
        }
    }

    return result;
}

function formatDecimal(value, tokens, hasFill = false) {
    // 分析整数和小数部分格式
    let intFormat = [], decFormat = [];
    let inDecimal = false;
    let useThousand = false;

    for (const tok of tokens) {
        if (tok.type === TokenType.DECIMAL) { inDecimal = true; continue; }
        if ([TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q].includes(tok.type)) {
            (inDecimal ? decFormat : intFormat).push(tok);
        }
        if (tok.type === TokenType.THOUSAND && !inDecimal) useThousand = true;
    }

    // 四舍五入到指定小数位
    const decPlaces = decFormat.length;
    let rounded = decPlaces > 0 ? value.toFixed(decPlaces) : Math.round(value).toString();
    let [intPart, decPart = ''] = rounded.split('.');

    // 整数部分格式化
    const minIntLen = intFormat.filter(t => t.type === TokenType.DIGIT_0).length || 1;
    intPart = intPart.padStart(minIntLen, '0');

    // 千分位
    if (useThousand) {
        intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }

    // 小数部分格式化
    if (decFormat.length > 0) {
        decPart = decPart.padEnd(decFormat.length, '0');
        // 处理 # 和 ? 的尾部零
        let trimEnd = 0;
        for (let i = decFormat.length - 1; i >= 0; i--) {
            if (decFormat[i].type === TokenType.DIGIT_HASH && decPart[i] === '0') trimEnd++;
            else break;
        }
        if (trimEnd > 0) decPart = decPart.slice(0, -trimEnd);
    }

    // 格式化数字字符串
    const numStr = decPart ? intPart + '.' + decPart : intPart;

    // 会计格式：返回结构化数据
    if (hasFill) {
        return formatAccountingStyle(tokens, numStr);
    }

    // 普通格式：组装结果字符串
    let result = '';
    for (const tok of tokens) {
        switch (tok.type) {
            case TokenType.DIGIT_0:
            case TokenType.DIGIT_HASH:
            case TokenType.DIGIT_Q:
                break; // 已处理
            case TokenType.DECIMAL:
                result += numStr;
                intPart = ''; decPart = '';
                break;
            case TokenType.THOUSAND:
                break; // 已处理
            case TokenType.PERCENT:
                result += '%';
                break;
            case TokenType.LITERAL:
                result += tok.value;
                break;
            case TokenType.SKIP:
                result += ' ';
                break;
            case TokenType.FILL:
                result += tok.value;
                break;
        }
    }

    // 如果没有小数点 token，处理整数格式（支持分散数字格式如 #"°"##"′"##"″"）
    if (!tokens.some(t => t.type === TokenType.DECIMAL)) {
        // 检测是否是分散数字格式（数字占位符之间有字面量）
        let hasLiteralBetweenDigits = false;
        let inDigits = false;
        for (let i = 0; i < tokens.length; i++) {
            const t = tokens[i].type;
            const isDigit = [TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q].includes(t);
            if (isDigit) {
                if (!inDigits && i > 0) {
                    // 检查前面是否有字面量且更前面有数字
                    for (let j = i - 1; j >= 0; j--) {
                        if (tokens[j].type === TokenType.LITERAL) {
                            for (let k = j - 1; k >= 0; k--) {
                                if ([TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q].includes(tokens[k].type)) {
                                    hasLiteralBetweenDigits = true;
                                    break;
                                }
                            }
                        }
                        if (hasLiteralBetweenDigits) break;
                    }
                }
                inDigits = true;
            } else if (tokens[i].type === TokenType.LITERAL) {
                inDigits = false;
            }
        }

        if (hasLiteralBetweenDigits) {
            // 分散数字格式：按 token 顺序处理，从右到左分配数字
            // 1. 收集数字占位符组（被字面量分隔）
            const segments = []; // {type: 'digits'|'literal', count?, value?}
            let currentDigits = 0;
            let hasZero = false; // 当前组是否有 0（需要显示前导零）

            for (const tok of tokens) {
                const isDigit = [TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q].includes(tok.type);
                if (isDigit) {
                    currentDigits++;
                    if (tok.type === TokenType.DIGIT_0) hasZero = true;
                } else if (tok.type === TokenType.THOUSAND) {
                    // 千分位跳过
                } else {
                    if (currentDigits > 0) {
                        segments.push({ type: 'digits', count: currentDigits, hasZero });
                        currentDigits = 0;
                        hasZero = false;
                    }
                    if (tok.type === TokenType.LITERAL) {
                        segments.push({ type: 'literal', value: tok.value });
                    } else if (tok.type === TokenType.SKIP) {
                        segments.push({ type: 'literal', value: ' ' });
                    } else if (tok.type === TokenType.PERCENT) {
                        segments.push({ type: 'literal', value: '%' });
                    }
                }
            }
            if (currentDigits > 0) {
                segments.push({ type: 'digits', count: currentDigits, hasZero });
            }

            // 2. 从右到左分配数字到各组
            const digitSegments = segments.filter(s => s.type === 'digits');
            let numStr = intPart;
            let pos = numStr.length;

            for (let i = digitSegments.length - 1; i >= 0; i--) {
                const seg = digitSegments[i];
                if (pos >= seg.count) {
                    seg.result = numStr.slice(pos - seg.count, pos);
                    pos -= seg.count;
                } else if (pos > 0) {
                    seg.result = numStr.slice(0, pos).padStart(seg.count, seg.hasZero ? '0' : '');
                    pos = 0;
                } else {
                    seg.result = seg.hasZero ? '0'.repeat(seg.count) : '';
                }
            }
            // 剩余数字加到第一个数字组
            if (pos > 0 && digitSegments.length > 0) {
                digitSegments[0].result = numStr.slice(0, pos) + digitSegments[0].result;
            }

            // 3. 组装结果
            result = '';
            let digitIdx = 0;
            for (const seg of segments) {
                if (seg.type === 'digits') {
                    result += digitSegments[digitIdx].result;
                    digitIdx++;
                } else {
                    result += seg.value;
                }
            }
        } else {
            // 普通格式：数字占位符连续，使用原有逻辑
            const numPos = tokens.findIndex(t => [TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q].includes(t.type));
            if (numPos >= 0) {
                let prefix = '', suffix = '';
                let passedNum = false;
                for (const tok of tokens) {
                    if ([TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q, TokenType.THOUSAND].includes(tok.type)) {
                        passedNum = true;
                        continue;
                    }
                    if (tok.type === TokenType.PERCENT) { suffix += '%'; continue; }
                    if (tok.type === TokenType.LITERAL) { if (passedNum) suffix += tok.value; else prefix += tok.value; continue; }
                    if (tok.type === TokenType.SKIP) { if (passedNum) suffix += ' '; else prefix += ' '; continue; }
                }
                result = prefix + intPart + suffix;
            }
        }
    }

    return result || intPart;
}

// 会计格式结构化数据处理
// 格式如: _ ¥* #,##0.00_ => { prefix: ' ¥', fill: ' ', number: '1,234.00', suffix: ' ' }
// 负数段: _ ¥* -#,##0.00_ => { prefix: ' ¥', fill: ' ', number: '-1,000.00', suffix: ' ' }
function formatAccountingStyle(tokens, numStr) {
    let prefix = '';
    let numberPrefix = ''; // FILL之后、数字之前的字面量（如负号）
    let suffix = '';
    let fillChar = ' ';
    let phase = 'prefix'; // prefix -> afterFill -> number -> suffix
    let seenDigit = false; // 是否已经遇到数字占位符

    for (const tok of tokens) {
        if (tok.type === TokenType.FILL) {
            fillChar = tok.value;
            phase = 'afterFill';
            continue;
        }

        const isDigitToken = [TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q, TokenType.DECIMAL, TokenType.THOUSAND].includes(tok.type);

        if (phase === 'prefix') {
            if (isDigitToken) {
                phase = 'number';
                seenDigit = true;
            } else if (tok.type === TokenType.LITERAL) {
                prefix += tok.value;
            } else if (tok.type === TokenType.SKIP) {
                prefix += ' ';
            }
        } else if (phase === 'afterFill') {
            // FILL 之后，数字之前的字面量（如负号）
            if (isDigitToken) {
                phase = 'number';
                seenDigit = true;
            } else if (tok.type === TokenType.LITERAL) {
                numberPrefix += tok.value;
            } else if (tok.type === TokenType.SKIP) {
                numberPrefix += ' ';
            }
        } else if (phase === 'number') {
            if (isDigitToken) {
                continue; // 数字部分已处理
            } else {
                phase = 'suffix';
                if (tok.type === TokenType.LITERAL) {
                    suffix += tok.value;
                } else if (tok.type === TokenType.SKIP) {
                    suffix += ' ';
                } else if (tok.type === TokenType.PERCENT) {
                    suffix += '%';
                }
            }
        } else if (phase === 'suffix') {
            if (tok.type === TokenType.LITERAL) {
                suffix += tok.value;
            } else if (tok.type === TokenType.SKIP) {
                suffix += ' ';
            } else if (tok.type === TokenType.PERCENT) {
                suffix += '%';
            }
        }
    }

    // 检查是否是零值显示模式（如 "-"??，用字面量替代零）
    // 如果 numberPrefix 是非空格字面量（如 "-"），且数字是零，则只显示字面量
    let finalNumber = numberPrefix + numStr;
    if (numberPrefix && numberPrefix.trim()) {
        const isZeroValue = /^0+(\.0+)?$/.test(numStr);
        if (isZeroValue) {
            // 用空格替换数字部分，保留字面量（如 "-"）
            const spaces = numStr.replace(/[0.]/g, ' ');
            finalNumber = numberPrefix + spaces;
        }
    }

    return {
        isAccounting: true,
        prefix,
        fill: fillChar,
        number: finalNumber,
        suffix
    };
}

function formatExponent(value, tokens) {
    // 获取指数符号
    const expToken = tokens.find(t => t.type === TokenType.EXPONENT);
    const expSign = expToken?.sign || '+';

    // 计算小数位数
    let decPlaces = 0;
    let inDecimal = false;
    for (const tok of tokens) {
        if (tok.type === TokenType.DECIMAL) inDecimal = true;
        else if (tok.type === TokenType.EXPONENT) break;
        else if (inDecimal && tok.type === TokenType.DIGIT_0) decPlaces++;
    }

    // 格式化
    let expStr = value.toExponential(decPlaces);
    let [mantissa, exp] = expStr.split('e');
    let expNum = parseInt(exp);

    // 构建结果
    let result = mantissa + 'E';
    if (expSign === '+' || expNum >= 0) result += expNum >= 0 ? '+' : '';
    result += expNum;

    return result;
}

// ==================== 分数格式 ====================

function formatFraction(value, tokens) {
    if (isNaN(value)) return String(value);

    const isNegative = value < 0;
    const absValue = Math.abs(value);
    const intPart = Math.floor(absValue);
    let fracPart = absValue - intPart;

    // 分析分数格式
    let hasIntPart = tokens.some(t => t.type === TokenType.DIGIT_HASH || t.type === TokenType.DIGIT_0);
    let fractionIdx = tokens.findIndex(t => t.type === TokenType.FRACTION);
    let denomTokens = tokens.slice(fractionIdx + 1).filter(t =>
        [TokenType.DIGIT_0, TokenType.DIGIT_HASH, TokenType.DIGIT_Q].includes(t.type)
    );

    // 固定分母判断
    let fixedDenom = 0;
    let denomDigits = denomTokens.length;
    if (denomTokens.every(t => t.type === TokenType.DIGIT_0)) {
        // 可能是固定分母如 /16
        const denomStr = tokens.slice(fractionIdx + 1)
            .filter(t => t.type === TokenType.LITERAL)
            .map(t => t.value).join('');
        if (/^\d+$/.test(denomStr)) fixedDenom = parseInt(denomStr);
    }

    // 转换为分数
    let numerator, denominator;
    if (fixedDenom > 0) {
        denominator = fixedDenom;
        numerator = Math.round(fracPart * denominator);
    } else {
        // 自动约分，限制分母位数
        const maxDenom = Math.pow(10, denomDigits) - 1;
        const frac = toFraction(fracPart, maxDenom);
        numerator = frac[0];
        denominator = frac[1];
    }

    // 约分
    if (numerator > 0) {
        const g = gcd(numerator, denominator);
        numerator /= g;
        denominator /= g;
    }

    // 格式化输出
    let result = '';
    if (intPart > 0 || numerator === 0) {
        result = String(intPart);
        if (numerator > 0) result += ' ';
    }
    if (numerator > 0) {
        const numStr = String(numerator).padStart(denomDigits, ' ');
        const denStr = String(denominator).padStart(denomDigits, ' ');
        result += `${numStr}/${denStr}`;
    }

    return (isNegative && result !== '0' ? '-' : '') + result.trim();
}

// 小数转分数（连分数算法）
function toFraction(decimal, maxDenom) {
    if (decimal === 0) return [0, 1];

    let h1 = 1, h2 = 0, k1 = 0, k2 = 1;
    let b = decimal;

    do {
        const a = Math.floor(b);
        let h = a * h1 + h2;
        let k = a * k1 + k2;
        if (k > maxDenom) break;
        h2 = h1; h1 = h;
        k2 = k1; k1 = k;
        if (b - a < 1e-10) break;
        b = 1 / (b - a);
    } while (true);

    return [h1, k1];
}

function gcd(a, b) {
    a = Math.abs(a);
    b = Math.abs(b);
    while (b) { const t = b; b = a % b; a = t; }
    return a;
}

// ==================== 日期字符串解析 ====================

export function dateStrToDate(input) {
    if (typeof input !== 'string') return null;

    const now = new Date();
    const currentYear = now.getFullYear();

    const patterns = [
        // yyyy-mm-dd hh:mm:ss 或 yyyy/mm/dd hh:mm:ss
        [/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?: (\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/, (y, m, d, h = 0, i = 0, s = 0) => {
            const hasTime = h || i || s;
            return {
                date: new Date(y, m - 1, d, h, i, s),
                numfmt: hasTime ? (s ? "yyyy/m/d h:mm:ss" : "yyyy/m/d h:mm") : "yyyy/m/d",
                type: hasTime ? "dateTime" : "date"
            };
        }],
        // yyyy年mm月dd日
        [/^(\d{4})年(\d{1,2})月(\d{1,2})日$/, (y, m, d) => ({
            date: new Date(y, m - 1, d),
            numfmt: 'yyyy"年"m"月"d"日"',
            type: "date"
        })],
        // yyyy年mm月
        [/^(\d{4})年(\d{1,2})月$/, (y, m) => ({
            date: new Date(y, m - 1, 1),
            numfmt: 'yyyy"年"m"月"',
            type: "date"
        })],
        // mm月dd日
        [/^(\d{1,2})月(\d{1,2})日$/, (m, d) => ({
            date: new Date(currentYear, m - 1, d),
            numfmt: 'm"月"d"日"',
            type: "date"
        })],
        // mm-dd 或 mm/dd
        [/^(\d{1,2})[-\/](\d{1,2})$/, (m, d) => ({
            date: new Date(currentYear, m - 1, d),
            numfmt: 'm/d',
            type: "date"
        })],
    ];

    for (const [regex, parser] of patterns) {
        const match = input.match(regex);
        if (match) {
            return parser(...match.slice(1).map(v => v === undefined ? v : +v));
        }
    }

    return null;
}

// ==================== 小数位增减 ====================

export function modifyDecimalPlaces(fmt, type) {
    if (!fmt) fmt = '0';

    // 查找小数点位置
    const decIdx = fmt.indexOf('.');
    if (type === 'add') {
        if (decIdx === -1) return fmt + '.0';
        // 在小数点后添加一个 0
        const before = fmt.slice(0, decIdx + 1);
        const after = fmt.slice(decIdx + 1);
        const zeroEnd = after.search(/[^0#?]/);
        if (zeroEnd === -1) return before + after + '0';
        return before + after.slice(0, zeroEnd) + '0' + after.slice(zeroEnd);
    } else {
        if (decIdx === -1) return fmt;
        const before = fmt.slice(0, decIdx + 1);
        const after = fmt.slice(decIdx + 1);
        // 找到小数部分的结尾
        const zeroEnd = after.search(/[^0#?]/);
        const decPart = zeroEnd === -1 ? after : after.slice(0, zeroEnd);
        const rest = zeroEnd === -1 ? '' : after.slice(zeroEnd);
        if (decPart.length <= 1) {
            // 移除小数点
            return fmt.slice(0, decIdx) + rest;
        }
        return before + decPart.slice(0, -1) + rest;
    }
}

// ==================== 双向日期转换（兼容原 helpers.js） ====================

export function dateTrans(input) {
    if (typeof input === 'number') {
        return serialToDate(input);
    } else if (input instanceof Date) {
        return dateToSerial(input);
    }
    return input;
}

// ==================== 导出 ====================

export default { formatValue, numFmtFun, dateStrToDate, dateToSerial, dateTrans, modifyDecimalPlaces };
