const MAX_OUTLINE_LEVEL = 7;

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

function _resolveAreas(sheet, range) {
    const fallback = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];
    const source = range ?? fallback;
    const areaList = Array.isArray(source) ? source : [source];
    const areas = [];

    areaList.forEach(item => {
        const parsed = typeof item === 'string' ? sheet.rangeStrToNum(item) : item;
        const normalized = _normalizeArea(parsed);
        if (normalized) areas.push(normalized);
    });

    return areas;
}

function _collectIndexes(areas, axis) {
    const startKey = axis === 'row' ? 'r' : 'c';
    const indexes = new Set();
    areas.forEach(area => {
        const start = area.s[startKey];
        const end = area.e[startKey];
        for (let i = start; i <= end; i++) {
            indexes.add(i);
        }
    });
    return Array.from(indexes).sort((a, b) => a - b);
}

function _normalizeMaxLevel(maxLevel) {
    const num = Number(maxLevel);
    if (!Number.isFinite(num)) return MAX_OUTLINE_LEVEL;
    return Math.max(1, Math.min(MAX_OUTLINE_LEVEL, Math.round(num)));
}

function _getTarget(sheet, axis, index) {
    return axis === 'row' ? sheet.getRow(index) : sheet.getCol(index);
}

function _snapshotState(sheet, axis, indexes) {
    return indexes.map(index => {
        const target = _getTarget(sheet, axis, index);
        return {
            index,
            outlineLevel: target._outlineLevel ?? 0,
            collapsed: target._collapsed ?? false
        };
    });
}

function _applyState(sheet, axis, states) {
    states.forEach(state => {
        const target = _getTarget(sheet, axis, state.index);
        target._outlineLevel = state.outlineLevel;
        target._collapsed = state.collapsed;
    });
}

function _changeOutline(sheet, range, options = {}, axis = 'row', delta = 1) {
    const areas = _resolveAreas(sheet, range);
    const indexes = _collectIndexes(areas, axis);
    const indexKey = axis === 'row' ? 'rows' : 'cols';
    const opName = `${delta > 0 ? 'group' : 'ungroup'}${axis === 'row' ? 'Rows' : 'Cols'}`;
    const eventSuffix = `${delta > 0 ? 'Group' : 'Ungroup'}${axis === 'row' ? 'Rows' : 'Cols'}`;
    const maxLevel = _normalizeMaxLevel(options.maxLevel);

    if (indexes.length === 0) {
        return { success: false, canceled: true, changedCount: 0, [`${axis}Count`]: 0 };
    }

    if (sheet.SN.Event.hasListeners(`before${eventSuffix}`)) {
        const beforeEvent = sheet.SN.Event.emit(`before${eventSuffix}`, {
            sheet,
            [indexKey]: indexes,
            areas,
            maxLevel
        });
        if (beforeEvent.canceled) {
            return { success: false, canceled: true, changedCount: 0, [`${axis}Count`]: indexes.length };
        }
    }

    const previous = _snapshotState(sheet, axis, indexes);
    let changedCount = 0;
    let limitedCount = 0;
    const next = previous.map(item => {
        const oldLevel = item.outlineLevel;
        let outlineLevel = Math.max(0, Math.min(maxLevel, oldLevel + delta));
        let collapsed = item.collapsed;
        if (outlineLevel === 0) collapsed = false;

        if (delta > 0 && oldLevel >= maxLevel) {
            limitedCount++;
        }
        if (outlineLevel !== oldLevel || collapsed !== item.collapsed) {
            changedCount++;
        }
        return { index: item.index, outlineLevel, collapsed };
    });

    const summary = {
        success: true,
        [indexKey]: indexes,
        [`${axis}Count`]: indexes.length,
        changedCount,
        maxLevel,
        limitedCount
    };

    if (changedCount === 0) return summary;

    return sheet.SN._withOperation(opName, {
        sheet,
        [indexKey]: indexes,
        changedCount,
        maxLevel
    }, () => {
        sheet.SN.UndoRedo.add({
            undo: () => _applyState(sheet, axis, previous),
            redo: () => _applyState(sheet, axis, next)
        });
        _applyState(sheet, axis, next);

        sheet.SN._recordChange({
            type: 'outline',
            action: opName,
            sheet,
            [indexKey]: indexes,
            changedCount,
            maxLevel
        });

        sheet.SN.Event.emit(`after${eventSuffix}`, {
            sheet,
            [indexKey]: indexes,
            changedCount,
            maxLevel,
            limitedCount
        });

        return summary;
    });
}

export function groupRows(range, options = {}) {
    return _changeOutline(this, range, options, 'row', 1);
}

export function ungroupRows(range, options = {}) {
    return _changeOutline(this, range, options, 'row', -1);
}

export function groupCols(range, options = {}) {
    return _changeOutline(this, range, options, 'col', 1);
}

export function ungroupCols(range, options = {}) {
    return _changeOutline(this, range, options, 'col', -1);
}
