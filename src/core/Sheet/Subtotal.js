const SUBTOTAL_FUNC_CODE = {
    average: 1,
    countNums: 2,
    count: 3,
    max: 4,
    min: 5,
    product: 6,
    stdev: 7,
    stdevp: 8,
    sum: 9,
    var: 10,
    varp: 11
};

function _normalizeArea(area) {
    if (!area?.s || !area?.e) return null;
    return {
        s: {
            r: Math.min(area.s.r, area.e.r),
            c: Math.min(area.s.c, area.e.c)
        },
        e: {
            r: Math.max(area.s.r, area.e.r),
            c: Math.max(area.s.c, area.e.c)
        }
    };
}

function _resolveArea(sheet, range) {
    const fallback = sheet.activeAreas?.[0] ?? { s: sheet.activeCell, e: sheet.activeCell };
    const raw = range ?? fallback;
    const parsed = typeof raw === 'string' ? sheet.rangeStrToNum(raw) : raw;
    const normalized = _normalizeArea(parsed);
    if (!normalized) throw new Error('无效的分类汇总区域');
    return normalized;
}

function _resolveColumnIndex(sheet, value) {
    if (Number.isInteger(value) && value >= 0) return value;
    if (typeof value === 'string') {
        const raw = value.replace(/\$/g, '').trim().toUpperCase();
        if (/^[A-Z]+$/.test(raw)) return sheet.Utils.charToNum(raw);
        if (/^\d+$/.test(raw)) return Number(raw);
    }
    return null;
}

function _resolveColumns(sheet, values) {
    const source = Array.isArray(values) ? values : [values];
    const cols = [];
    source.forEach(item => {
        const idx = _resolveColumnIndex(sheet, item);
        if (Number.isInteger(idx) && idx >= 0) cols.push(idx);
    });
    return Array.from(new Set(cols)).sort((a, b) => a - b);
}

function _normalizeFunction(funcName) {
    const key = String(funcName ?? 'sum').trim();
    if (!key) return 'sum';
    if (SUBTOTAL_FUNC_CODE[key]) return key;
    const lower = key.toLowerCase();
    if (SUBTOTAL_FUNC_CODE[lower]) return lower;
    if (lower === 'avg') return 'average';
    if (lower === 'counta') return 'count';
    if (lower === 'countnum' || lower === 'countnums') return 'countNums';
    if (lower === 'stddev') return 'stdev';
    if (lower === 'stddevp') return 'stdevp';
    return 'sum';
}

function _buildGroupKey(cell) {
    const val = cell.calcVal;
    if (val === null || val === undefined || val === '') return '__blank__';
    const type = typeof val;
    if (type === 'number' && Number.isNaN(val)) return 'number:NaN';
    return `${type}:${String(val)}`;
}

function _buildGroupLabel(cell) {
    const show = cell.showVal;
    if (show === null || show === undefined || show === '') return '空白';
    return String(show);
}

function _cellRef(sheet, r, c) {
    return sheet.Utils.cellNumToStr({ r, c });
}

function _isSubtotalFormula(value) {
    return typeof value === 'string' && /^\s*=SUBTOTAL\s*\(/i.test(value);
}

function _rowHasSubtotalFormula(sheet, rowIndex, columns) {
    return columns.some(col => _isSubtotalFormula(sheet.getCell(rowIndex, col).editVal));
}

function _getFormulaCode(funcName) {
    return SUBTOTAL_FUNC_CODE[_normalizeFunction(funcName)] ?? 9;
}

function _normalizeOptions(sheet, area, options = {}) {
    const hasHeaders = options.hasHeaders !== false;
    const summaryBelowData = options.summaryBelowData !== false;
    const replaceExisting = options.replaceExisting !== false;
    const withGrandTotal = options.withGrandTotal !== false;
    const functionName = _normalizeFunction(options.function ?? options.useFunction ?? 'sum');
    const formulaCode = _getFormulaCode(functionName);

    const groupByRaw = options.groupBy ?? options.atEachChangeIn ?? area.s.c;
    const groupBy = _resolveColumnIndex(sheet, groupByRaw);
    if (!Number.isInteger(groupBy)) {
        throw new Error('分类字段无效');
    }

    const subtotalColsRaw = options.subtotalColumns ?? options.addSubtotalTo ?? options.columns;
    let subtotalColumns = _resolveColumns(sheet, subtotalColsRaw);
    if (subtotalColumns.length === 0) {
        for (let c = area.s.c; c <= area.e.c; c++) {
            if (c === groupBy) continue;
            subtotalColumns.push(c);
        }
    }

    if (subtotalColumns.length === 0) {
        throw new Error('请至少选择一个汇总列');
    }

    const outOfRange = [groupBy, ...subtotalColumns].some(c => c < area.s.c || c > area.e.c);
    if (outOfRange) {
        throw new Error('分类字段或汇总列必须在选区内');
    }

    return {
        hasHeaders,
        summaryBelowData,
        replaceExisting,
        withGrandTotal,
        functionName,
        formulaCode,
        groupBy,
        subtotalColumns
    };
}

function _collectGroups(sheet, startRow, endRow, groupBy) {
    const groups = [];
    let currentStart = startRow;
    let currentKey = _buildGroupKey(sheet.getCell(startRow, groupBy));
    let currentLabel = _buildGroupLabel(sheet.getCell(startRow, groupBy));

    for (let r = startRow + 1; r <= endRow; r++) {
        const key = _buildGroupKey(sheet.getCell(r, groupBy));
        if (key === currentKey) continue;

        groups.push({ start: currentStart, end: r - 1, label: currentLabel });
        currentStart = r;
        currentKey = key;
        currentLabel = _buildGroupLabel(sheet.getCell(r, groupBy));
    }
    groups.push({ start: currentStart, end: endRow, label: currentLabel });
    return groups;
}

function _applySubtotalStyle(sheet, rowIndex, colIndex) {
    const cell = sheet.getCell(rowIndex, colIndex);
    cell.font = { ...(cell.font || {}), bold: true };
}

export function subtotal(range, options = {}) {
    const area = _resolveArea(this, range);
    const config = _normalizeOptions(this, area, options);
    if (this.areaHaveMerge(area)) {
        throw new Error('分类汇总不支持包含合并单元格的区域');
    }

    const dataStart = config.hasHeaders ? area.s.r + 1 : area.s.r;
    if (dataStart > area.e.r) {
        throw new Error('分类汇总至少需要一行数据');
    }

    const beforeEvent = this.SN.Event.emit('beforeSubtotal', {
        sheet: this,
        range: area,
        options: config
    });
    if (beforeEvent.canceled) {
        return { success: false, canceled: true };
    }

    return this.SN._withOperation('subtotal', {
        sheet: this,
        range: area,
        options: config
    }, () => {
        const previousOutlinePr = this.outlinePr ? { ...this.outlinePr } : null;
        const nextOutlinePr = {
            ...(this.outlinePr || {}),
            summaryBelow: config.summaryBelowData
        };
        this.SN.UndoRedo.add({
            undo: () => { this.outlinePr = previousOutlinePr ? { ...previousOutlinePr } : previousOutlinePr; },
            redo: () => { this.outlinePr = { ...nextOutlinePr }; }
        });
        this.outlinePr = { ...nextOutlinePr };

        let mutableEnd = area.e.r;
        let removedCount = 0;
        if (config.replaceExisting) {
            for (let r = mutableEnd; r >= dataStart; r--) {
                if (!_rowHasSubtotalFormula(this, r, config.subtotalColumns)) continue;
                this.delRows(r, 1);
                removedCount++;
            }
            mutableEnd -= removedCount;
        }

        if (dataStart > mutableEnd) {
            throw new Error('分类汇总后没有可处理的数据');
        }

        for (let r = dataStart; r <= mutableEnd; r++) {
            const row = this.getRow(r);
            row._outlineLevel = 0;
            row._collapsed = false;
        }

        const groups = _collectGroups(this, dataStart, mutableEnd, config.groupBy);
        let insertedSubtotalRows = 0;

        for (let i = groups.length - 1; i >= 0; i--) {
            const group = groups[i];
            const insertAt = config.summaryBelowData ? group.end + 1 : group.start;
            this.addRows(insertAt, 1);
            insertedSubtotalRows++;

            const label = `${group.label} 小计`;
            this.getCell(insertAt, config.groupBy).editVal = label;
            _applySubtotalStyle(this, insertAt, config.groupBy);

            config.subtotalColumns.forEach(col => {
                if (col === config.groupBy) return;
                const formula = `=SUBTOTAL(${config.formulaCode},${_cellRef(this, group.start, col)}:${_cellRef(this, group.end, col)})`;
                this.getCell(insertAt, col).editVal = formula;
                _applySubtotalStyle(this, insertAt, col);
            });

            const detailStart = config.summaryBelowData ? group.start : group.start + 1;
            const detailEnd = config.summaryBelowData ? group.end : group.end + 1;
            for (let r = detailStart; r <= detailEnd; r++) {
                const row = this.getRow(r);
                row._outlineLevel = Math.min(7, (row._outlineLevel || 0) + 1);
                row._collapsed = false;
            }
        }

        let grandTotalRow = null;
        if (config.withGrandTotal) {
            const grandInsertAt = mutableEnd + insertedSubtotalRows + 1;
            this.addRows(grandInsertAt, 1);
            grandTotalRow = grandInsertAt;
            this.getCell(grandInsertAt, config.groupBy).editVal = '总计';
            _applySubtotalStyle(this, grandInsertAt, config.groupBy);
            config.subtotalColumns.forEach(col => {
                if (col === config.groupBy) return;
                const formula = `=SUBTOTAL(${config.formulaCode},${_cellRef(this, dataStart, col)}:${_cellRef(this, grandInsertAt - 1, col)})`;
                this.getCell(grandInsertAt, col).editVal = formula;
                _applySubtotalStyle(this, grandInsertAt, col);
            });
        }

        const resultRange = {
            s: { ...area.s },
            e: {
                r: mutableEnd + insertedSubtotalRows + (config.withGrandTotal ? 1 : 0),
                c: area.e.c
            }
        };
        const summary = {
            success: true,
            range: area,
            resultRange,
            groupCount: groups.length,
            removedCount,
            insertedSubtotalRows,
            grandTotalRow
        };

        this.SN._recordChange({
            type: 'subtotal',
            sheet: this,
            range: area,
            resultRange,
            groupCount: groups.length,
            removedCount,
            insertedSubtotalRows,
            grandTotalRow,
            options: config
        });

        this.SN.Event.emit('afterSubtotal', {
            sheet: this,
            ...summary,
            options: config
        });

        return summary;
    });
}
