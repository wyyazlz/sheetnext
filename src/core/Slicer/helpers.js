export function isBlankValue(value) {
    return value === null || value === undefined || value === '';
}

export function buildSlicerKey(value) {
    if (isBlankValue(value)) return 'm:';
    if (value instanceof Date) return `d:${value.toISOString()}`;
    const valueType = typeof value;
    if (valueType === 'number') return `n:${value}`;
    if (valueType === 'boolean') return `b:${value ? 1 : 0}`;
    return `s:${String(value)}`;
}

export function parseSlicerKey(key) {
    if (!key || key.length < 2) return null;
    const type = key[0];
    const raw = key.slice(2);
    switch (type) {
        case 'm':
            return null;
        case 'n': {
            const num = Number(raw);
            return Number.isNaN(num) ? null : num;
        }
        case 'b':
            return raw === '1';
        case 'd':
            return raw ? new Date(raw) : null;
        case 's':
            return raw;
        default:
            return raw;
    }
}

export function formatSlicerValue(value) {
    if (isBlankValue(value)) return '(blank)';
    if (value instanceof Date) return value.toISOString();
    return String(value);
}
