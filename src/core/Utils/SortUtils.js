/** @param {*} a @param {*} b @param {{order?:string, nulls?:string, locale?:string, customOrder?:Array, caseSensitive?:boolean}} [options] @returns {number} Compare sort values. */
export function compare(a, b, options = {}) {
    const { order = 'asc', nulls = 'last', locale = 'zh-CN', customOrder, caseSensitive = false } = options;
    const mul = order === 'asc' ? 1 : -1;

    if (customOrder?.length) {
        const iA = customOrder.indexOf(a);
        const iB = customOrder.indexOf(b);
        if (iA !== -1 && iB !== -1) return (iA - iB) * mul;
        if (iA !== -1) return -1;
        if (iB !== -1) return 1;
    }

    const emptyA = a === null || a === undefined || a === '';
    const emptyB = b === null || b === undefined || b === '';
    if (emptyA && emptyB) return 0;
    if (emptyA) return nulls === 'last' ? 1 : -1;
    if (emptyB) return nulls === 'last' ? -1 : 1;

    const nA = typeof a === 'number' ? a : parseFloat(a);
    const nB = typeof b === 'number' ? b : parseFloat(b);
    if (!isNaN(nA) && !isNaN(nB)) {
        return (nA - nB) * mul;
    }

    const strA = String(a);
    const strB = String(b);

    if (caseSensitive) {
        if (strA < strB) return -1 * mul;
        if (strA > strB) return 1 * mul;
        return 0;
    }

    return strA.localeCompare(strB, locale, { sensitivity: 'base' }) * mul;
}

/** @param {Sheet} sheet @param {RangeRef|string} range @param {Array|Object} sortKeys @param {{hasHeader?:boolean, caseSensitive?:boolean}} [options] @returns {boolean} Sort rows in range. */
export function sortRange(sheet, range, sortKeys, options = {}) {
    const SN = sheet.SN;
    const { hasHeader = false, caseSensitive = false } = options;

    const r = typeof range === 'string' ? sheet.rangeStrToNum(range) : range;
    const startRow = r.s.r;
    const endRow = r.e.r;
    const startCol = r.s.c;
    const endCol = r.e.c;
    const dataStartRow = hasHeader ? startRow + 1 : startRow;

    if (startRow > endRow) return false;

    if (sheet.areaHaveMerge({ s: { r: startRow, c: startCol }, e: { r: endRow, c: endCol } })) {
        SN.Utils.toast(SN.t('core.utils.sortUtils.toast.msg001'));
        return false;
    }

    const keys = (Array.isArray(sortKeys) ? sortKeys : [sortKeys]).map(key => ({
        ...key,
        col: typeof key.col === 'string' ? SN.Utils.charToNum(key.col) : key.col
    }));

    const original = [];
    for (let rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
        const row = sheet.getRow(rowIndex);
        original.push(row?.cells?.slice() || []);
    }

    const rows = [];
    for (let rowIndex = dataStartRow; rowIndex <= endRow; rowIndex++) {
        const row = sheet.getRow(rowIndex);
        rows.push({
            idx: rowIndex,
            cells: row?.cells?.slice() || [],
            values: keys.map(key => sheet.getCell(rowIndex, key.col)?.showVal)
        });
    }

    rows.sort((rowA, rowB) => {
        for (let i = 0; i < keys.length; i++) {
            const result = compare(rowA.values[i], rowB.values[i], {
                order: keys[i].order || 'asc',
                customOrder: keys[i].customOrder,
                caseSensitive
            });
            if (result !== 0) return result;
        }
        return 0;
    });

    const sorted = hasHeader ? [original[0], ...rows.map(row => row.cells)] : rows.map(row => row.cells);
    for (let i = 0; i < rows.length; i++) {
        const row = sheet.getRow(dataStartRow + i);
        if (!row) continue;
        row.cells = rows[i].cells;
        row.cells.forEach((cell, cellIndex) => {
            if (!cell) return;
            cell.row = row;
            cell.cIndex = cellIndex;
        });
    }

    SN.UndoRedo.add({
        undo: () => {
            for (let i = 0; i < original.length; i++) {
                const row = sheet.getRow(startRow + i);
                if (!row) continue;
                row.cells = original[i];
                row.cells.forEach((cell, cellIndex) => {
                    if (!cell) return;
                    cell.row = row;
                    cell.cIndex = cellIndex;
                });
            }
            SN._r();
        },
        redo: () => {
            for (let i = 0; i < sorted.length; i++) {
                const row = sheet.getRow(startRow + i);
                if (!row) continue;
                row.cells = sorted[i];
                row.cells.forEach((cell, cellIndex) => {
                    if (!cell) return;
                    cell.row = row;
                    cell.cIndex = cellIndex;
                });
            }
            SN._r();
        }
    });

    return true;
}
