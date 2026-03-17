function _cloneCell(cell) {
    if (!cell || !Number.isFinite(cell.r) || !Number.isFinite(cell.c)) return null;
    return { r: cell.r, c: cell.c };
}

export function cloneRange(range) {
    if (!range?.s || !range?.e) return null;
    return {
        s: _cloneCell(range.s),
        e: _cloneCell(range.e)
    };
}

function _getAxisKey(axis) {
    return axis === 'row' ? 'r' : 'c';
}

function _parseCellToken(token, utils) {
    const match = /^(\$?)([A-Z]+)(\$?)(\d+)$/i.exec(String(token || ''));
    if (!match) return null;
    return {
        colAbs: match[1] === '$',
        c: utils.charToNum(match[2].toUpperCase()),
        rowAbs: match[3] === '$',
        r: parseInt(match[4], 10) - 1
    };
}

function _formatCellToken(cell, utils) {
    return `${cell.colAbs ? '$' : ''}${utils.numToChar(cell.c)}${cell.rowAbs ? '$' : ''}${cell.r + 1}`;
}

function _parseColToken(token, utils) {
    const match = /^(\$?)([A-Z]+)$/i.exec(String(token || ''));
    if (!match) return null;
    return {
        abs: match[1] === '$',
        c: utils.charToNum(match[2].toUpperCase())
    };
}

function _formatColToken(col, utils) {
    return `${col.abs ? '$' : ''}${utils.numToChar(col.c)}`;
}

function _parseRowToken(token) {
    const match = /^(\$?)(\d+)$/.exec(String(token || ''));
    if (!match) return null;
    return {
        abs: match[1] === '$',
        r: parseInt(match[2], 10) - 1
    };
}

function _formatRowToken(row) {
    return `${row.abs ? '$' : ''}${row.r + 1}`;
}

function _normalizeSheetName(sheetName) {
    if (typeof sheetName !== 'string') return '';
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
        return sheetName.slice(1, -1).replace(/''/g, "'");
    }
    return sheetName;
}

function _shiftToken(token, utils, axis, index, count, mode) {
    if (typeof token !== 'string' || !token) return token;

    if (token.includes(':')) {
        const [startToken, endToken] = token.split(':');
        const startCell = _parseCellToken(startToken, utils);
        const endCell = _parseCellToken(endToken, utils);
        if (startCell && endCell) {
            const nextRange = shiftRange({
                s: { r: startCell.r, c: startCell.c },
                e: { r: endCell.r, c: endCell.c }
            }, axis, index, count, mode);
            if (!nextRange) return '#REF!';

            const nextStart = { ...startCell, ...nextRange.s };
            const nextEnd = { ...endCell, ...nextRange.e };
            const startText = _formatCellToken(nextStart, utils);
            const endText = _formatCellToken(nextEnd, utils);
            return startText === endText ? startText : `${startText}:${endText}`;
        }

        const startCol = _parseColToken(startToken, utils);
        const endCol = _parseColToken(endToken, utils);
        if (startCol && endCol) {
            if (axis !== 'col') return token;
            const nextRange = _shiftAxisRange({
                s: startCol.c,
                e: endCol.c
            }, index, count, mode);
            if (!nextRange) return '#REF!';
            const nextStart = { ...startCol, c: nextRange.s };
            const nextEnd = { ...endCol, c: nextRange.e };
            return `${_formatColToken(nextStart, utils)}:${_formatColToken(nextEnd, utils)}`;
        }

        const startRow = _parseRowToken(startToken);
        const endRow = _parseRowToken(endToken);
        if (startRow && endRow) {
            if (axis !== 'row') return token;
            const nextRange = _shiftAxisRange({
                s: startRow.r,
                e: endRow.r
            }, index, count, mode);
            if (!nextRange) return '#REF!';
            const nextStart = { ...startRow, r: nextRange.s };
            const nextEnd = { ...endRow, r: nextRange.e };
            return `${_formatRowToken(nextStart)}:${_formatRowToken(nextEnd)}`;
        }

        return token;
    }

    const cell = _parseCellToken(token, utils);
    if (!cell) return token;

    const nextCell = shiftCell({ r: cell.r, c: cell.c }, axis, index, count, mode);
    if (!nextCell) return '#REF!';
    return _formatCellToken({ ...cell, ...nextCell }, utils);
}

function _shiftAxisRange(range, index, count, mode) {
    if (!range || !Number.isFinite(range.s) || !Number.isFinite(range.e)) return null;
    const next = { s: range.s, e: range.e };
    if (mode === 'insert') {
        if (index <= next.s) {
            next.s += count;
            next.e += count;
        } else if (index <= next.e + 1) {
            next.e += count;
        }
        return next;
    }

    const delEnd = index + count - 1;
    if (next.e < index) return next;
    if (next.s > delEnd) {
        next.s -= count;
        next.e -= count;
        return next;
    }

    const nextStart = next.s < index ? next.s : (next.e > delEnd ? index : null);
    const nextEnd = next.e > delEnd ? next.e - count : (next.s < index ? index - 1 : null);
    if (nextStart === null || nextEnd === null || nextStart > nextEnd) {
        return null;
    }

    next.s = nextStart;
    next.e = nextEnd;
    return next;
}

export function shiftCell(cell, axis, index, count, mode, options = {}) {
    const next = _cloneCell(cell);
    if (!next) return null;

    const key = _getAxisKey(axis);
    const pos = next[key];
    if (mode === 'insert') {
        if (pos >= index) next[key] += count;
        return next;
    }

    const delEnd = index + count - 1;
    if (pos < index) return next;
    if (pos > delEnd) {
        next[key] -= count;
        return next;
    }

    if (options.onDelete === 'clamp') {
        next[key] = index;
        return next;
    }

    return null;
}

export function shiftRange(range, axis, index, count, mode) {
    const next = cloneRange(range);
    if (!next) return null;

    const key = _getAxisKey(axis);
    if (mode === 'insert') {
        if (index <= next.s[key]) {
            next.s[key] += count;
            next.e[key] += count;
        } else if (index <= next.e[key] + 1) {
            next.e[key] += count;
        }
        return next;
    }

    const delEnd = index + count - 1;
    if (next.e[key] < index) return next;
    if (next.s[key] > delEnd) {
        next.s[key] -= count;
        next.e[key] -= count;
        return next;
    }

    const nextStart = next.s[key] < index ? next.s[key] : (next.e[key] > delEnd ? index : null);
    const nextEnd = next.e[key] > delEnd ? next.e[key] - count : (next.s[key] < index ? index - 1 : null);
    if (nextStart === null || nextEnd === null || nextStart > nextEnd) {
        return null;
    }

    next.s[key] = nextStart;
    next.e[key] = nextEnd;
    return next;
}

export function shiftAreaRef(areaRef, utils, axis, index, count, mode) {
    const next = _shiftToken(String(areaRef || '').trim(), utils, axis, index, count, mode);
    return next === '#REF!' ? null : next;
}

export function shiftSqref(sqref, utils, axis, index, count, mode) {
    if (!sqref) return '';
    return String(sqref)
        .split(/\s+/)
        .map(part => shiftAreaRef(part, utils, axis, index, count, mode))
        .filter(Boolean)
        .join(' ');
}

function _replaceFormulaSegment(segment, callback) {
    const pattern = /(^|[^A-Za-z0-9_\]])((?:'(?:[^']|'')+'|[A-Za-z0-9_.]+)!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+:\$?[A-Z]+|\$?\d+:\$?\d+)/g;
    return segment.replace(pattern, (match, boundary, sheetPrefix = '', token) => {
        const next = callback(sheetPrefix, token);
        if (next === null) return match;
        if (next === '#REF!') return `${boundary}#REF!`;
        return `${boundary}${sheetPrefix}${next}`;
    });
}

export function shiftFormulaRefs(formula, currentSheetName, changedSheetName, utils, axis, index, count, mode) {
    if (typeof formula !== 'string' || !formula) return formula;
    if (!changedSheetName) return formula;

    let result = '';
    let cursor = 0;

    while (cursor < formula.length) {
        if (formula[cursor] === '"') {
            let end = cursor + 1;
            while (end < formula.length) {
                if (formula[end] !== '"') {
                    end++;
                    continue;
                }
                if (formula[end + 1] === '"') {
                    end += 2;
                    continue;
                }
                end++;
                break;
            }
            result += formula.slice(cursor, end);
            cursor = end;
            continue;
        }

        let end = cursor;
        while (end < formula.length && formula[end] !== '"') {
            end++;
        }
        const segment = formula.slice(cursor, end);
        result += _replaceFormulaSegment(segment, (sheetPrefix, token) => {
            const targetSheetName = sheetPrefix
                ? _normalizeSheetName(sheetPrefix.slice(0, -1))
                : currentSheetName;

            if (targetSheetName !== changedSheetName) {
                return null;
            }

            return _shiftToken(token, utils, axis, index, count, mode);
        });
        cursor = end;
    }

    return result;
}
