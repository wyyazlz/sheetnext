const NUMERIC_FUNCTIONS = new Set(['sum', 'average', 'max', 'min', 'product', 'countNums', 'stdev', 'stdevp', 'var', 'varp']);

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

function _resolveRanges(sheet, ranges) {
    const fallback = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];
    const source = ranges ?? fallback;
    const list = Array.isArray(source) ? source : [source];
    const result = [];

    list.forEach(item => {
        const parsed = typeof item === 'string' ? sheet.rangeStrToNum(item) : item;
        const normalized = _normalizeArea(parsed);
        if (normalized) result.push(normalized);
    });

    return result;
}

function _resolveDestination(sheet, destination, fallbackRange) {
    if (!destination) return { ...fallbackRange.s };
    if (typeof destination === 'string') {
        const cell = sheet.Utils.cellStrToNum(destination.replace(/\$/g, '').trim());
        if (!cell || cell.r < 0 || cell.c < 0) throw new Error('合并计算目标单元格无效');
        return { r: cell.r, c: cell.c };
    }
    if (typeof destination === 'object' && destination.r >= 0 && destination.c >= 0) {
        return { r: destination.r, c: destination.c };
    }
    throw new Error('合并计算目标单元格无效');
}

function _normalizeFunction(fn) {
    const raw = String(fn ?? 'sum').trim();
    if (!raw) return 'sum';
    const lower = raw.toLowerCase();
    if (lower === 'avg') return 'average';
    if (lower === 'countnum' || lower === 'countnums') return 'countNums';
    if (lower === 'stddev') return 'stdev';
    if (lower === 'stddevp') return 'stdevp';
    if (lower === 'variance') return 'var';
    if (lower === 'variancep') return 'varp';
    return lower;
}

function _toNumber(val) {
    if (typeof val === 'number' && Number.isFinite(val)) return val;
    if (typeof val === 'string') {
        const text = val.trim();
        if (!text) return null;
        const num = Number(text);
        if (Number.isFinite(num)) return num;
    }
    return null;
}

function _isBlank(val) {
    return val === '' || val === null || val === undefined;
}

function _compute(funcName, items) {
    if (funcName === 'count') {
        return items.filter(item => !_isBlank(item.raw)).length;
    }

    const nums = items
        .map(item => item.num)
        .filter(num => Number.isFinite(num));

    switch (funcName) {
        case 'sum':
            return nums.length ? nums.reduce((total, num) => total + num, 0) : '';
        case 'average':
            return nums.length ? nums.reduce((total, num) => total + num, 0) / nums.length : '';
        case 'max':
            return nums.length ? Math.max(...nums) : '';
        case 'min':
            return nums.length ? Math.min(...nums) : '';
        case 'product':
            return nums.length ? nums.reduce((total, num) => total * num, 1) : '';
        case 'countnums':
        case 'countNums':
            return nums.length;
        case 'stdev': {
            if (nums.length <= 1) return '';
            const mean = nums.reduce((total, num) => total + num, 0) / nums.length;
            const variance = nums.reduce((total, num) => total + Math.pow(num - mean, 2), 0) / (nums.length - 1);
            return Math.sqrt(variance);
        }
        case 'stdevp': {
            if (nums.length === 0) return '';
            const mean = nums.reduce((total, num) => total + num, 0) / nums.length;
            const variance = nums.reduce((total, num) => total + Math.pow(num - mean, 2), 0) / nums.length;
            return Math.sqrt(variance);
        }
        case 'var': {
            if (nums.length <= 1) return '';
            const mean = nums.reduce((total, num) => total + num, 0) / nums.length;
            return nums.reduce((total, num) => total + Math.pow(num - mean, 2), 0) / (nums.length - 1);
        }
        case 'varp': {
            if (nums.length === 0) return '';
            const mean = nums.reduce((total, num) => total + num, 0) / nums.length;
            return nums.reduce((total, num) => total + Math.pow(num - mean, 2), 0) / nums.length;
        }
        default:
            throw new Error('不支持的合并计算函数');
    }
}

export function consolidate(ranges, options = {}) {
    if (ranges && typeof ranges === 'object' && !Array.isArray(ranges) && !ranges.s) {
        options = ranges;
        ranges = options.sourceRanges ?? options.ranges;
    }

    const sourceRanges = _resolveRanges(this, ranges ?? options.sourceRanges ?? options.ranges);
    if (sourceRanges.length === 0) throw new Error('请至少提供一个源区域');
    if (sourceRanges.some(range => this.areaHaveMerge(range))) {
        throw new Error('合并计算不支持包含合并单元格的源区域');
    }

    const funcName = _normalizeFunction(options.function ?? options.func ?? 'sum');
    const skipBlanks = options.skipBlanks !== false;

    const rowCount = Math.max(...sourceRanges.map(range => range.e.r - range.s.r + 1));
    const colCount = Math.max(...sourceRanges.map(range => range.e.c - range.s.c + 1));
    const destination = _resolveDestination(this, options.destination, sourceRanges[0]);
    const targetRange = {
        s: { ...destination },
        e: { r: destination.r + rowCount - 1, c: destination.c + colCount - 1 }
    };

    const beforeEvent = this.SN.Event.emit('beforeConsolidate', {
        sheet: this,
        sourceRanges,
        targetRange,
        functionName: funcName,
        skipBlanks
    });
    if (beforeEvent.canceled) return { success: false, canceled: true };

    return this.SN._withOperation('consolidate', {
        sheet: this,
        sourceRanges,
        targetRange,
        functionName: funcName
    }, () => {
        let writtenCount = 0;
        for (let rOffset = 0; rOffset < rowCount; rOffset++) {
            for (let cOffset = 0; cOffset < colCount; cOffset++) {
                const items = [];
                sourceRanges.forEach(range => {
                    const sourceRow = range.s.r + rOffset;
                    const sourceCol = range.s.c + cOffset;
                    if (sourceRow > range.e.r || sourceCol > range.e.c) return;

                    const cell = this.getCell(sourceRow, sourceCol);
                    const raw = cell.calcVal;
                    if (skipBlanks && _isBlank(raw)) return;

                    const num = _toNumber(raw);
                    if (NUMERIC_FUNCTIONS.has(funcName) && funcName !== 'count' && num === null) return;
                    items.push({ raw, num });
                });

                const result = _compute(funcName, items);
                const targetCell = this.getCell(destination.r + rOffset, destination.c + cOffset);
                targetCell.editVal = result === '' ? '' : result;
                if (result !== '') writtenCount++;
            }
        }

        const summary = {
            success: true,
            sourceRanges,
            targetRange,
            functionName: funcName,
            rowCount,
            colCount,
            writtenCount
        };

        this.SN._recordChange({
            type: 'consolidate',
            sheet: this,
            sourceRanges,
            targetRange,
            functionName: funcName,
            rowCount,
            colCount,
            writtenCount
        });

        this.SN.Event.emit('afterConsolidate', {
            sheet: this,
            ...summary
        });

        return summary;
    });
}
