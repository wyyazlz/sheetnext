const _STYLE_FIELDS = ['font', 'fill', 'alignment', 'border', 'numFmt', 'protection'];
const _MERGEABLE_FIELDS = new Set(['font', 'fill', 'alignment', 'border']);

function _isPlainObject(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
}

function _normalizeStyleValue(value) {
    if (value === undefined || value === null || value === '') return undefined;

    if (Array.isArray(value)) {
        const list = value
            .map(item => _normalizeStyleValue(item))
            .filter(item => item !== undefined);
        return list.length ? list : undefined;
    }

    if (_isPlainObject(value)) {
        const result = {};
        Object.keys(value).sort().forEach(key => {
            const normalized = _normalizeStyleValue(value[key]);
            if (normalized !== undefined) result[key] = normalized;
        });
        return Object.keys(result).length ? result : undefined;
    }

    return value;
}

/** @param {any} value @returns {any} */
export function _cloneStyleValue(value) {
    if (value === undefined) return undefined;
    return JSON.parse(JSON.stringify(value));
}

/** @param {Object} style @returns {Object} */
export function _pruneStyleObject(style) {
    const normalized = _normalizeStyleValue(style);
    return _isPlainObject(normalized) ? normalized : {};
}

/** @param {any} a @param {any} b @returns {boolean} */
export function _isSameStyleValue(a, b) {
    const left = _normalizeStyleValue(a);
    const right = _normalizeStyleValue(b);
    if (left === undefined && right === undefined) return true;
    return JSON.stringify(left) === JSON.stringify(right);
}

/** @param {Object|null} oldStyle @param {Object|null} newStyle @returns {string[]} */
export function _getChangedStyleFields(oldStyle, newStyle) {
    return _STYLE_FIELDS.filter(field => !_isSameStyleValue(oldStyle?.[field], newStyle?.[field]));
}

/** @param {string} field @param {Object|null} rowStyle @param {Object|null} colStyle @returns {any} */
export function _getInheritedStyleField(field, rowStyle, colStyle) {
    const rowValue = _normalizeStyleValue(rowStyle?.[field]);
    const colValue = _normalizeStyleValue(colStyle?.[field]);

    if (rowValue === undefined && colValue === undefined) return undefined;

    if (_MERGEABLE_FIELDS.has(field) && _isPlainObject(rowValue) && _isPlainObject(colValue)) {
        return {
            ..._cloneStyleValue(colValue),
            ..._cloneStyleValue(rowValue)
        };
    }

    return _cloneStyleValue(rowValue ?? colValue);
}

/** @param {Cell} cell @returns {Object|null} */
export function _getCellExplicitStyle(cell) {
    if (!cell) return null;
    if (cell._style !== undefined) return cell._style;
    if (cell._xmlObj?.['_$s']) return cell.style;
    return null;
}

/** @param {Cell} cell @param {Object} [options={}] @returns {Object} */
export function _getCellEffectiveStyle(cell, options = {}) {
    const explicitStyle = options.cellStyle ?? _getCellExplicitStyle(cell);
    const rowStyle = options.rowStyle ?? cell.row?._rowStyle;
    const colStyle = options.colStyle ?? cell.row?.sheet?.cols?.[cell.cIndex]?._colStyle;
    const result = {};

    _STYLE_FIELDS.forEach(field => {
        const explicitValue = _normalizeStyleValue(explicitStyle?.[field]);
        const nextValue = explicitValue !== undefined
            ? _cloneStyleValue(explicitValue)
            : _getInheritedStyleField(field, rowStyle, colStyle);
        if (nextValue !== undefined) result[field] = nextValue;
    });

    return _pruneStyleObject(result);
}
