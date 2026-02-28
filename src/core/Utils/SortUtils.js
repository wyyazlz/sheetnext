export function compare(a, b, options = {}) {
    const { order = 'asc', nulls = 'last', locale = 'zh-CN', customOrder, caseSensitive = false } = options;
    const mul = order === 'asc' ? 1 : -1;

    // 自定义顺序
    if (customOrder?.length) {
        const iA = customOrder.indexOf(a), iB = customOrder.indexOf(b);
        if (iA !== -1 && iB !== -1) return (iA - iB) * mul;
        if (iA !== -1) return -1;
        if (iB !== -1) return 1;
    }

    // 空值处理
    const emptyA = a === null || a === undefined || a === '';
    const emptyB = b === null || b === undefined || b === '';
    if (emptyA && emptyB) return 0;
    if (emptyA) return nulls === 'last' ? 1 : -1;
    if (emptyB) return nulls === 'last' ? -1 : 1;

    // 优先尝试数字比较
    const nA = typeof a === 'number' ? a : parseFloat(a);
    const nB = typeof b === 'number' ? b : parseFloat(b);
    if (!isNaN(nA) && !isNaN(nB)) {
        return (nA - nB) * mul;
    }

    // 字符串比较
    let strA = String(a);
    let strB = String(b);

    if (caseSensitive) {
        if (strA < strB) return -1 * mul;
        if (strA > strB) return 1 * mul;
        return 0;
    }

    return strA.localeCompare(strB, locale, { sensitivity: 'base' }) * mul;
}

export function sortRange(sheet, range, sortKeys) {
    const SN = sheet.SN;

    // 范围处理
    const r = typeof range === 'string' ? sheet.rangeStrToNum(range) : range;
    const startRow = r.s.r;
    const endRow = r.e.r;
    const startCol = r.s.c;
    const endCol = r.e.c;

    if (startRow > endRow) return false;

    // 合并单元格检查
    if (sheet.areaHaveMerge({ s: { r: startRow, c: startCol }, e: { r: endRow, c: endCol } })) {
        SN.Utils.toast(SN.t('core.utils.sortUtils.toast.msg001'));
        return false;
    }

    // 标准化排序键（支持 'A' 或 0）
    const keys = (Array.isArray(sortKeys) ? sortKeys : [sortKeys]).map(k => ({
        ...k,
        col: typeof k.col === 'string' ? SN.Utils.charToNum(k.col) : k.col
    }));

    // 提取行数据
    const rows = [];
    for (let ri = startRow; ri <= endRow; ri++) {
        const row = sheet.getRow(ri);
        rows.push({
            idx: ri,
            cells: row?.cells?.slice() || [],
            values: keys.map(k => sheet.getCell(ri, k.col)?.showVal)
        });
    }

    // 保存原始顺序
    const original = rows.map(r => r.cells.slice());

    // 多级排序
    rows.sort((a, b) => {
        for (let i = 0; i < keys.length; i++) {
            const result = compare(a.values[i], b.values[i], {
                order: keys[i].order || 'asc',
                customOrder: keys[i].customOrder
            });
            if (result !== 0) return result;
        }
        return 0;
    });

    // 应用排序
    const sorted = rows.map(r => r.cells);
    for (let i = 0; i < rows.length; i++) {
        const row = sheet.getRow(startRow + i);
        if (row) {
            row.cells = sorted[i];
            row.cells.forEach((cell, ci) => {
                if (cell) {
                    cell.row = row;
                    cell.cIndex = ci;
                }
            });
        }
    }

    // 撤销/重做
    SN.UndoRedo.add({
        undo: () => {
            for (let i = 0; i < original.length; i++) {
                const row = sheet.getRow(startRow + i);
                if (row) {
                    row.cells = original[i];
                    row.cells.forEach((cell, ci) => {
                        if (cell) { cell.row = row; cell.cIndex = ci; }
                    });
                }
            }
            SN._r();
        },
        redo: () => {
            for (let i = 0; i < sorted.length; i++) {
                const row = sheet.getRow(startRow + i);
                if (row) {
                    row.cells = sorted[i];
                    row.cells.forEach((cell, ci) => {
                        if (cell) { cell.row = row; cell.cIndex = ci; }
                    });
                }
            }
            SN._r();
        }
    });

    return true;
}
