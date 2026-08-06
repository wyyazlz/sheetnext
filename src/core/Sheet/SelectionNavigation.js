const DIRECTION_INFO = {
    up: { axis: 'r', step: -1 },
    down: { axis: 'r', step: 1 },
    left: { axis: 'c', step: -1 },
    right: { axis: 'c', step: 1 }
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

function _sameCell(a, b) {
    return a?.r === b?.r && a?.c === b?.c;
}

function _sameArea(a, b) {
    return !!a && !!b && _sameCell(a.s, b.s) && _sameCell(a.e, b.e);
}

function _findMergeAt(sheet, cell) {
    return sheet.merges.find(merge =>
        cell.r >= merge.s.r && cell.r <= merge.e.r &&
        cell.c >= merge.s.c && cell.c <= merge.e.c
    );
}

function _getCellSpan(sheet, cell) {
    return _findMergeAt(sheet, cell) || { s: { ...cell }, e: { ...cell } };
}

function _normalizeFocus(sheet, cell) {
    const merge = _findMergeAt(sheet, cell);
    return merge ? { ...merge.s } : { ...cell };
}

function _getFocusArea(sheet, anchor, focus) {
    return _normalizeArea(sheet._getExpandMaxArea({
        s: { ...anchor },
        e: { ...focus }
    }));
}

function _inferAxisFocus(anchor, start, end, directionStep) {
    if (start === end) return start;
    if (anchor <= start) return end;
    if (anchor >= end) return start;
    if (directionStep < 0) return start;
    if (directionStep > 0) return end;
    return anchor;
}

function _resolveFocus(sheet, anchor, area, direction) {
    const stored = sheet._selectionFocus;
    if (stored && _sameArea(_getFocusArea(sheet, anchor, stored), area)) {
        return { ...stored };
    }

    const info = DIRECTION_INFO[direction];
    const rowStep = info.axis === 'r' ? info.step : 0;
    const colStep = info.axis === 'c' ? info.step : 0;
    return _normalizeFocus(sheet, {
        r: _inferAxisFocus(anchor.r, area.s.r, area.e.r, rowStep),
        c: _inferAxisFocus(anchor.c, area.s.c, area.e.c, colStep)
    });
}

function _moveFocusOne(sheet, focus, direction) {
    const span = _getCellSpan(sheet, focus);
    let target = { ...focus };

    if (direction === 'up') target.r = sheet._findBeforeShowRow(span.s.r, 1);
    else if (direction === 'down') target.r = sheet._findNextShowRow(span.e.r, 1);
    else if (direction === 'left') target.c = sheet._findBeforeShowCol(span.s.c, 1);
    else target.c = sheet._findNextShowCol(span.e.c, 1);

    return _normalizeFocus(sheet, target);
}

function _cellHasNavigationContent(cell) {
    if (!cell) return false;
    const value = cell._editVal;
    return (value !== undefined && value !== null && value !== '') ||
        !!cell._richText?.length || !!cell._spillParent;
}

function _pendingCellHasNavigationContent(cell) {
    if (!cell) return false;
    if (cell.f !== undefined && cell.f !== null) return true;
    if (cell.v !== undefined && cell.v !== null && cell.v !== '') return true;
    return cell.is !== undefined && cell.is !== null && cell.is !== '';
}

function _getPendingCellColumn(sheet, cell, fallbackIndex) {
    const match = String(cell?.['_$r'] || '').match(/[A-Z]+/i);
    return match ? sheet.Utils.charToNum(match[0]) : fallbackIndex;
}

function _getPendingCell(sheet, r, c) {
    const row = sheet._pendingXlsxRowMap?.get(r);
    if (!row) return null;
    const cells = Array.isArray(row.c) ? row.c : (row.c ? [row.c] : []);
    return cells.find((cell, index) => _getPendingCellColumn(sheet, cell, index) === c) || null;
}

function _cellAtHasNavigationContent(sheet, r, c) {
    const cell = sheet.rows[r]?.cells?.[c];
    if (cell) return _cellHasNavigationContent(cell);
    return _pendingCellHasNavigationContent(_getPendingCell(sheet, r, c));
}

function _isAxisIndexVisible(sheet, axis, index) {
    if (axis === 'r') {
        const row = sheet.rows[index];
        const pending = !row ? sheet._getPendingXlsxRowState?.(index) : null;
        return !(row?.hidden || pending?.hidden || pending?.filterHidden);
    }
    return !sheet.cols[index]?.hidden;
}

function _addPendingRowPositions(sheet, positions, axis, fixed, row, r) {
    const cells = Array.isArray(row?.c) ? row.c : (row?.c ? [row.c] : []);
    cells.forEach((cell, index) => {
        if (!_pendingCellHasNavigationContent(cell)) return;
        const c = _getPendingCellColumn(sheet, cell, index);
        if (axis === 'r' && c === fixed) positions.add(r);
        if (axis === 'c' && r === fixed) positions.add(c);
    });
}

function _addMergePositions(sheet, positions, axis, fixed) {
    sheet.merges.forEach(merge => {
        if (!_cellAtHasNavigationContent(sheet, merge.s.r, merge.s.c)) return;
        if (axis === 'r' && fixed >= merge.s.c && fixed <= merge.e.c) {
            for (let r = merge.s.r; r <= merge.e.r; r++) positions.add(r);
        } else if (axis === 'c' && fixed >= merge.s.r && fixed <= merge.e.r) {
            for (let c = merge.s.c; c <= merge.e.c; c++) positions.add(c);
        }
    });
}

function _getOccupiedAxisPositions(sheet, axis, fixed, limit) {
    const positions = new Set();
    if (axis === 'r') {
        Object.keys(sheet.rows).forEach(key => {
            const r = Number(key);
            if (_cellHasNavigationContent(sheet.rows[r]?.cells?.[fixed])) positions.add(r);
        });
        sheet._pendingXlsxRowMap?.forEach((row, r) => {
            _addPendingRowPositions(sheet, positions, axis, fixed, row, r);
        });
    } else {
        const row = sheet.rows[fixed];
        Object.keys(row?.cells || {}).forEach(key => {
            const c = Number(key);
            if (_cellHasNavigationContent(row.cells[c])) positions.add(c);
        });
        const pendingRow = sheet._pendingXlsxRowMap?.get(fixed);
        if (pendingRow) _addPendingRowPositions(sheet, positions, axis, fixed, pendingRow, fixed);
    }

    _addMergePositions(sheet, positions, axis, fixed);
    return [...positions]
        .filter(index => index >= 0 && index <= limit && _isAxisIndexVisible(sheet, axis, index))
        .sort((a, b) => a - b);
}

function _getAxisLimit(sheet, axis) {
    return Math.max(0, (axis === 'r' ? sheet.rowCount : sheet.colCount) - 1);
}

function _findNextVisibleIndex(sheet, axis, index, step, limit) {
    for (let current = index + step; current >= 0 && current <= limit; current += step) {
        if (_isAxisIndexVisible(sheet, axis, current)) return current;
    }
    return index;
}

function _findJumpIndex(sheet, focus, direction) {
    const { axis, step } = DIRECTION_INFO[direction];
    const fixed = axis === 'r' ? focus.c : focus.r;
    const current = focus[axis];
    const limit = _getAxisLimit(sheet, axis);
    const occupied = _getOccupiedAxisPositions(sheet, axis, fixed, limit);
    const occupiedSet = new Set(occupied);
    const next = _findNextVisibleIndex(sheet, axis, current, step, limit);
    if (next === current) return current;

    if (occupiedSet.has(current) && occupiedSet.has(next)) {
        let target = next;
        while (true) {
            const following = _findNextVisibleIndex(sheet, axis, target, step, limit);
            if (following === target || !occupiedSet.has(following)) break;
            target = following;
        }
        return target;
    }

    if (step > 0) {
        const target = occupied.find(index => index > current);
        if (target !== undefined) return target;
    } else {
        for (let i = occupied.length - 1; i >= 0; i--) {
            if (occupied[i] < current) return occupied[i];
        }
    }

    let boundary = step > 0 ? limit : 0;
    while (boundary !== current && !_isAxisIndexVisible(sheet, axis, boundary)) boundary -= step;
    return boundary;
}

function _moveFocusToBoundary(sheet, focus, direction) {
    const { axis, step } = DIRECTION_INFO[direction];
    const span = _getCellSpan(sheet, focus);
    const edge = step < 0 ? span.s[axis] : span.e[axis];
    const start = { ...focus, [axis]: edge };
    const target = { ...focus, [axis]: _findJumpIndex(sheet, start, direction) };
    return _normalizeFocus(sheet, target);
}

function _ensureTargetExists(sheet, target) {
    if (target.r >= sheet.rowCount) {
        sheet._virtualRowCount = Math.min(sheet._maxVirtualRowCount, target.r + 1);
        sheet._clearSizeCache();
    }
    if (target.c >= sheet.colCount) {
        sheet._virtualColCount = Math.min(sheet._maxVirtualColCount, target.c + 1);
        sheet._clearSizeCache();
    }
}

function _canSelectTarget(sheet, target) {
    const protection = sheet.protection;
    if (!protection?.enabled) return true;
    const cell = sheet.getCell(target.r, target.c);
    if (protection.isCellSelectable(cell)) return true;

    const key = protection.isCellLocked(cell)
        ? 'core.sheet.sheetProtection.message.cannotSelectLocked'
        : 'core.sheet.sheetProtection.message.cannotSelectUnlocked';
    sheet.SN?.Utils?.toast?.(sheet.SN.t(key));
    return false;
}

/**
 * Move or extend the current selection using Excel-style navigation.
 * @param {'up'|'down'|'left'|'right'} direction
 * @param {{extend?:boolean,jump?:boolean}} [options]
 * @returns {CellNum|null}
 */
export function moveSelection(direction, options = {}) {
    if (!DIRECTION_INFO[direction]) return null;

    const anchor = _normalizeFocus(this, this.activeCell);
    const areas = this.activeAreas;
    const currentArea = _normalizeArea(areas[areas.length - 1]);
    if (!currentArea) return null;

    const focus = _resolveFocus(this, anchor, currentArea, direction);
    const target = options.jump
        ? _moveFocusToBoundary(this, focus, direction)
        : _moveFocusOne(this, focus, direction);
    if (_sameCell(target, focus) || !_canSelectTarget(this, target)) return null;

    _ensureTargetExists(this, target);
    if (!options.extend) {
        this.activeCell = target;
        if (!_sameCell(this.activeCell, target)) return null;
        this.activeAreas = [];
        return target;
    }

    const nextArea = _getFocusArea(this, anchor, target);
    const nextAreas = areas.slice(0, -1).concat(nextArea);
    this.activeAreas = nextAreas;
    if (!_sameArea(_normalizeArea(this.activeAreas[this.activeAreas.length - 1]), nextArea)) return null;
    this._selectionFocus = target;
    return target;
}
