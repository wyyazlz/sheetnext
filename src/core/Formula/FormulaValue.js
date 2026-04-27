export const FORMULA_ERROR_CODES = new Set([
    '#NULL!',
    '#DIV/0!',
    '#VALUE!',
    '#REF!',
    '#NAME?',
    '#NUM!',
    '#N/A',
    '#GETTING_DATA',
    '#SPILL!',
    '#CALC!'
]);

export class FormulaError extends Error {
    constructor(code = '#VALUE!') {
        super(normalizeErrorCode(code));
        this.name = 'FormulaError';
        this.code = this.message;
    }
}

export function normalizeErrorCode(value, fallback = '#VALUE!') {
    const text = value instanceof Error ? value.message : String(value ?? '');
    const match = text.match(/#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!)/);
    return match && FORMULA_ERROR_CODES.has(match[0]) ? match[0] : fallback;
}

export function isErrorCodeString(value) {
    return typeof value === 'string' && FORMULA_ERROR_CODES.has(value);
}

export function isErrorValue(value) {
    return value instanceof Error;
}

export function toFormulaError(value, fallback = '#VALUE!') {
    if (value instanceof FormulaError) return value;
    return new FormulaError(normalizeErrorCode(value, fallback));
}

export function errorCodeOf(value) {
    if (!isErrorValue(value)) return null;
    return normalizeErrorCode(value);
}

export function isBlankValue(value) {
    return value === null || value === undefined || value === '';
}

export function normalizeBlank(value) {
    return value === null || value === undefined ? '' : value;
}

export function throwIfError(value) {
    if (isErrorValue(value)) throw toFormulaError(value);
}

export function coerceNumber(value) {
    throwIfError(value);
    if (typeof value === 'number') return value;
    if (typeof value === 'boolean') return value ? 1 : 0;
    if (isBlankValue(value)) return 0;
    if (typeof value === 'string') {
        const text = value.trim();
        if (text === '') return 0;
        const num = Number(text);
        if (Number.isFinite(num)) return num;
    }
    throw new FormulaError('#VALUE!');
}

export function coerceLogical(value) {
    throwIfError(value);
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    if (isBlankValue(value)) return false;
    if (typeof value === 'string') {
        const text = value.trim().toUpperCase();
        if (text === 'TRUE') return true;
        if (text === 'FALSE') return false;
    }
    throw new FormulaError('#VALUE!');
}

export function coerceText(value) {
    throwIfError(value);
    if (isBlankValue(value)) return '';
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    return String(value);
}

export function firstErrorValue(...values) {
    for (const value of values) {
        if (Array.isArray(value)) {
            const nested = firstErrorValue(...value.flat(Infinity));
            if (nested) return nested;
        } else if (isErrorValue(value)) {
            return value;
        }
    }
    return null;
}

export function errorTypeNumber(value) {
    const code = errorCodeOf(value);
    if (!code) return null;
    return {
        '#NULL!': 1,
        '#DIV/0!': 2,
        '#VALUE!': 3,
        '#REF!': 4,
        '#NAME?': 5,
        '#NUM!': 6,
        '#N/A': 7,
        '#GETTING_DATA': 8,
        '#SPILL!': 9,
        '#CALC!': 14
    }[code] ?? null;
}
