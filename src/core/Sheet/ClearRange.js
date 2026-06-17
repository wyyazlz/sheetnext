const CLEAR_RANGE_TYPES = {
    all: {
        contents: true,
        formats: true,
        hyperlinks: true,
        dataValidations: true,
        comments: true,
        controls: true
    },
    contents: { contents: true },
    content: { contents: true },
    value: { contents: true },
    values: { contents: true },
    formats: { formats: true },
    format: { formats: true },
    style: { formats: true },
    styles: { formats: true },
    border: { borders: true },
    borders: { borders: true },
    hyperlinks: { hyperlinks: true },
    hyperlink: { hyperlinks: true },
    dataValidations: { dataValidations: true },
    dataValidation: { dataValidations: true },
    validations: { dataValidations: true },
    validation: { dataValidations: true },
    comments: { comments: true },
    comment: { comments: true },
    controls: { controls: true },
    control: { controls: true },
    merges: { merges: true },
    merge: { merges: true }
};

function _clone(value) {
    if (value === undefined || value === null) return value;
    return JSON.parse(JSON.stringify(value));
}

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

function _parseRangeString(sheet, ref) {
    const text = String(ref || '').replace(/['$]/g, '').trim().toUpperCase();
    if (!text) return null;
    if (/^\d+$/.test(text)) return sheet.rangeStrToNum(`${text}:${text}`);
    if (/^[A-Z]+$/.test(text)) return sheet.rangeStrToNum(`${text}:${text}`);
    return sheet.rangeStrToNum(text);
}

function _resolveAreas(sheet, range) {
    const fallback = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];
    const source = range ?? fallback;
    const list = Array.isArray(source) ? source : [source];
    const areas = [];
    list.forEach(item => {
        const parsed = typeof item === 'string' ? _parseRangeString(sheet, item) : item;
        const normalized = _normalizeArea(parsed);
        if (normalized) areas.push(normalized);
    });
    return areas;
}

function _normalizeClearOptions(options) {
    if (options === undefined || options === null) return { contents: true, formats: true };
    if (typeof options === 'string') return { ...(CLEAR_RANGE_TYPES[options] || { contents: true, formats: true }) };
    if (typeof options !== 'object') return { contents: true, formats: true };

    const normalized = {
        contents: !!(options.contents ?? options.content ?? options.values ?? options.value),
        formats: !!(options.formats ?? options.format ?? options.styles ?? options.style),
        borders: !!(options.borders ?? options.border),
        hyperlinks: !!(options.hyperlinks ?? options.hyperlink),
        dataValidations: !!(options.dataValidations ?? options.validations ?? options.validation),
        comments: !!(options.comments ?? options.comment),
        controls: !!(options.controls ?? options.control),
        merges: !!(options.merges ?? options.merge)
    };

    if (options.all === true) {
        normalized.contents = true;
        normalized.formats = true;
        normalized.hyperlinks = true;
        normalized.dataValidations = true;
        normalized.comments = true;
        normalized.controls = true;
        normalized.merges = !!(options.merges ?? options.merge);
    }
    return normalized;
}

function _hasClearOption(options) {
    return Object.values(options).some(Boolean);
}

function _cellHasClearableState(cell, options) {
    if (!cell) return false;
    if (options.contents && _hasContentState(cell)) return true;
    if (_shouldMaterializeStyle(options) && (cell._style !== undefined || cell._xmlObj?.['_$s'])) return true;
    if (options.hyperlinks && cell._hyperlink) return true;
    if (options.dataValidations && cell._dataValidation) return true;
    return false;
}

function _snapshotCell(cell) {
    return {
        cell,
        style: _clone(cell._style),
        editVal: cell._editVal,
        calcVal: cell._calcVal,
        showVal: cell._showVal,
        fmtResult: cell._fmtResult,
        type: cell._type,
        richText: _clone(cell._richText),
        hyperlink: _clone(cell._hyperlink),
        dataValidation: _clone(cell._dataValidation),
        spillRange: _clone(cell._spillRange),
        spillArray: _clone(cell._spillArray),
        spillParent: cell._spillParent,
        spillOffset: _clone(cell._spillOffset),
        isVolatile: cell._isVolatile
    };
}

function _restoreCell(snapshot) {
    snapshot.cell._style = _clone(snapshot.style);
    snapshot.cell._editVal = snapshot.editVal;
    snapshot.cell._calcVal = snapshot.calcVal;
    snapshot.cell._showVal = snapshot.showVal;
    snapshot.cell._fmtResult = snapshot.fmtResult;
    snapshot.cell._type = snapshot.type;
    snapshot.cell._richText = _clone(snapshot.richText);
    snapshot.cell._hyperlink = _clone(snapshot.hyperlink);
    snapshot.cell._dataValidation = _clone(snapshot.dataValidation);
    snapshot.cell._spillRange = _clone(snapshot.spillRange);
    snapshot.cell._spillArray = _clone(snapshot.spillArray);
    snapshot.cell._spillParent = snapshot.spillParent || null;
    snapshot.cell._spillOffset = _clone(snapshot.spillOffset);
    snapshot.cell._isVolatile = !!snapshot.isVolatile;
    _syncVolatileStatus(snapshot.cell, snapshot.isVolatile);
}

function _sameSnapshot(a, b) {
    return JSON.stringify({
        style: a.style,
        editVal: a.editVal,
        calcVal: a.calcVal,
        showVal: a.showVal,
        fmtResult: a.fmtResult,
        type: a.type,
        richText: a.richText,
        hyperlink: a.hyperlink,
        dataValidation: a.dataValidation,
        spillRange: a.spillRange,
        spillArray: a.spillArray,
        spillParent: a.spillParent ? `${a.spillParent.row.rIndex},${a.spillParent.cIndex}` : null,
        spillOffset: a.spillOffset,
        isVolatile: a.isVolatile
    }) === JSON.stringify({
        style: b.style,
        editVal: b.editVal,
        calcVal: b.calcVal,
        showVal: b.showVal,
        fmtResult: b.fmtResult,
        type: b.type,
        richText: b.richText,
        hyperlink: b.hyperlink,
        dataValidation: b.dataValidation,
        spillRange: b.spillRange,
        spillArray: b.spillArray,
        spillParent: b.spillParent ? `${b.spillParent.row.rIndex},${b.spillParent.cIndex}` : null,
        spillOffset: b.spillOffset,
        isVolatile: b.isVolatile
    });
}

function _shouldMaterializeStyle(options) {
    return !!(options.formats || options.borders || options.controls);
}

function _shouldNotifyValueChange(options) {
    return !!options.contents;
}

function _hasContentState(cell) {
    return cell._editVal !== undefined && cell._editVal !== ''
        || cell._calcVal !== undefined
        || cell._showVal !== undefined
        || cell._fmtResult !== undefined
        || cell._type !== undefined
        || !!cell._richText
        || !!cell._spillRange
        || !!cell._spillParent
        || !!cell._spillArray
        || !!cell._isVolatile;
}

function _syncVolatileStatus(cell, enabled) {
    const graph = cell._SN?.DependencyGraph;
    if (!graph) return;
    const sheetName = cell.row.sheet.name;
    if (enabled) graph.registerVolatile?.(sheetName, cell.row.rIndex, cell.cIndex);
    else graph.unregisterVolatile?.(sheetName, cell.row.rIndex, cell.cIndex);
}

function _areaCellCount(area) {
    return (area.e.r - area.s.r + 1) * (area.e.c - area.s.c + 1);
}

function _shouldUseSparseCellScan(sheet, areas) {
    const maxDirectCells = 100000;
    return areas.some(area => _areaCellCount(area) > maxDirectCells || (
        area.s.c === 0 && area.e.c === sheet.colCount - 1
    ) || (
        area.s.r === 0 && area.e.r === sheet.rowCount - 1
    ));
}

function _getPendingRowsInAreas(sheet, areas) {
    const rows = new Set();
    const indexes = sheet._pendingXlsxRowIndexes?.length
        ? sheet._pendingXlsxRowIndexes
        : [...(sheet._pendingXlsxRowMap?.keys?.() || [])];

    indexes.forEach(r => {
        if (areas.some(area => r >= area.s.r && r <= area.e.r)) rows.add(r);
    });
    return rows;
}

function _forEachClearableCell(sheet, areas, options, callback) {
    const visited = new Set();

    if (!_shouldUseSparseCellScan(sheet, areas)) {
        sheet.eachCells(areas, (r, c) => {
            const key = `${r},${c}`;
            if (visited.has(key)) return;
            visited.add(key);
            callback(sheet.getCell(r, c));
        });
        return;
    }

    const rowIndexes = new Set();
    sheet.rows.forEach((row, r) => {
        if (row && areas.some(area => r >= area.s.r && r <= area.e.r)) rowIndexes.add(r);
    });
    _getPendingRowsInAreas(sheet, areas).forEach(r => rowIndexes.add(r));

    rowIndexes.forEach(r => {
        const row = sheet.rows[r] || sheet._materializePendingRow?.(r);
        if (!row) return;
        row.cells.forEach((cell, c) => {
            if (!cell) return;
            const area = areas.find(item => r >= item.s.r && r <= item.e.r && c >= item.s.c && c <= item.e.c);
            if (!area) return;
            const key = `${r},${c}`;
            if (visited.has(key)) return;
            visited.add(key);
            if (_cellHasClearableState(cell, options)) callback(cell);
        });
    });
}

function _clearCell(cell, options) {
    if (options.contents) {
        if (_hasContentState(cell)) {
            if (cell._spillParent) cell._spillParent._clearSpillRefs?.();
            cell._clearVolatileStatus?.();
            cell._clearSpillRefs?.();
            cell._editVal = '';
            cell._calcVal = undefined;
            cell._showVal = undefined;
            cell._fmtResult = undefined;
            cell._type = undefined;
            cell._richText = null;
            cell._spillParent = null;
            cell._spillOffset = null;
            cell._spillArray = null;
        }
    }

    if (options.formats) {
        cell._style = {};
        cell._showVal = undefined;
        cell._fmtResult = undefined;
    } else if (options.borders && cell._style?.border) {
        delete cell._style.border;
    }

    if (options.controls && cell._style?.control) delete cell._style.control;
    if (options.hyperlinks) cell._hyperlink = null;
    if (options.dataValidations) cell._dataValidation = null;
}

function _rangeIntersects(a, b) {
    return !(a.s.c > b.e.c || a.e.c < b.s.c || a.s.r > b.e.r || a.e.r < b.s.r);
}

function _cellInAreas(sheet, cellRef, areas) {
    const cell = sheet.Utils.cellStrToNum(cellRef);
    return areas.some(area =>
        cell.r >= area.s.r && cell.r <= area.e.r &&
        cell.c >= area.s.c && cell.c <= area.e.c
    );
}

function _snapshotComments(sheet, areas) {
    if (!sheet.Comment?.getAll) return [];
    return sheet.Comment.getAll()
        .filter(comment => _cellInAreas(sheet, comment.cellRef, areas))
        .map(comment => ({
            cellRef: comment.cellRef,
            author: comment.author,
            text: comment.text,
            visible: comment.visible,
            width: comment.width,
            height: comment.height,
            marginLeft: comment.marginLeft,
            marginTop: comment.marginTop
        }));
}

function _clearComments(sheet, comments) {
    comments.forEach(item => sheet.Comment.remove(item.cellRef));
}

function _restoreComments(sheet, comments) {
    comments.forEach(item => sheet.Comment.add({ ...item }));
}

function _styleChanged(before, after) {
    return JSON.stringify(before) !== JSON.stringify(after);
}

function _clearStyle(style, options) {
    if (options.formats) return null;

    const next = _clone(style ?? {});
    if (options.borders) delete next.border;
    if (options.controls) delete next.control;
    return Object.keys(next).length ? next : null;
}

function _addRowsForStyleClear(sheet, area, rows) {
    const count = area.e.r - area.s.r + 1;
    if (count <= 100000) {
        for (let r = area.s.r; r <= area.e.r; r++) rows.add(r);
        return;
    }

    sheet._styledRows?.forEach(r => {
        if (r >= area.s.r && r <= area.e.r) rows.add(r);
    });
    sheet._pendingXlsxRowStateMap?.forEach((state, r) => {
        if (state?.rowStyle && r >= area.s.r && r <= area.e.r) rows.add(r);
    });
}

function _addColsForStyleClear(area, cols) {
    for (let c = area.s.c; c <= area.e.c; c++) cols.add(c);
}

function _collectRowColStyleChanges(sheet, areas, options) {
    if (!options.formats && !options.borders && !options.controls) return [];

    const changes = [];
    const rows = new Set();
    const cols = new Set();

    areas.forEach(area => {
        const isFullRow = area.s.c === 0 && area.e.c === sheet.colCount - 1;
        const isFullCol = area.s.r === 0 && area.e.r === sheet.rowCount - 1;

        if (isFullRow) _addRowsForStyleClear(sheet, area, rows);
        if (isFullCol) _addColsForStyleClear(area, cols);
    });

    rows.forEach(r => {
        const pendingState = !sheet.rows[r] ? sheet._pendingXlsxRowStateMap?.get(r) : null;
        const row = pendingState ? null : sheet.getRow(r);
        const before = _clone(pendingState ? pendingState.rowStyle : row._rowStyle);
        const after = _clearStyle(before, options);
        if (!_styleChanged(before, after)) return;
        if (pendingState) {
            changes.push({ type: 'pendingRow', index: r, state: pendingState, before, after });
            pendingState.rowStyle = after;
        } else {
            changes.push({ type: 'row', row, before, after });
            row._rowStyle = after;
        }
    });

    cols.forEach(c => {
        const col = sheet.getCol(c);
        const before = _clone(col._colStyle);
        const after = _clearStyle(before, options);
        if (!_styleChanged(before, after)) return;
        changes.push({ type: 'col', col, before, after });
        col._colStyle = after;
    });

    return changes;
}

function _applyRowColStyleChanges(changes, key) {
    changes.forEach(change => {
        if (change.type === 'row') change.row._rowStyle = _clone(change[key]);
        else if (change.type === 'col') change.col._colStyle = _clone(change[key]);
        else change.state.rowStyle = _clone(change[key]);
    });
}

function _getMergeChanges(sheet, areas, options) {
    if (!options.merges) return null;
    const before = _clone(sheet.merges);
    const removed = before.filter(merge => areas.some(area => _rangeIntersects(merge, area)));
    if (!removed.length) return null;
    const after = before.filter(merge => !removed.some(item => item.s.r === merge.s.r && item.s.c === merge.s.c));
    return { before, after, removed };
}

function _applyMerges(sheet, merges) {
    const affected = [];
    const seen = new Set();

    sheet.merges.forEach(merge => affected.push(merge));
    merges.forEach(merge => affected.push(merge));

    affected.forEach(area => {
        sheet.eachCells(area, (r, c) => {
            const key = `${r},${c}`;
            if (seen.has(key)) return;
            seen.add(key);
            const cell = sheet.getCell(r, c);
            cell.isMerged = false;
            cell.master = null;
        });
    });

    sheet.merges = _clone(merges) || [];
    sheet.merges.forEach(area => {
        sheet.eachCells(area, (r, c) => {
            const cell = sheet.getCell(r, c);
            cell.isMerged = true;
            cell.master = { r: area.s.r, c: area.s.c };
        });
    });
}

function _notifyAfterClear(sheet, changes, options) {
    sheet.Sparkline?.clearAllCache?.();
    sheet.CF?.clearAllCache?.();
    sheet.Drawing?._clearChartRefCaches?.();
    if (!sheet.SN.DependencyGraph) return;

    if (changes?.length && _shouldNotifyValueChange(options)) {
        changes.forEach(change => {
            const cell = change.after.cell;
            if (change.before.editVal !== change.after.editVal) {
                sheet.SN.DependencyGraph.notifyChange?.(sheet.name, cell.row.rIndex, cell.cIndex);
            }
            if ((typeof change.before.editVal === 'string' && change.before.editVal.startsWith('='))
                || (typeof change.after.editVal === 'string' && change.after.editVal.startsWith('='))) {
                sheet.SN.DependencyGraph.updateDependencies?.(cell);
            }
        });
    }

    if (sheet.SN.calcMode !== 'manual') {
        sheet.SN.DependencyGraph.recalculateSheetFormulas?.(sheet);
    }
}

/**
 * Clear cells in a range. Defaults to clearing contents and formats.
 * @param {RangeRef|RangeRef[]} range
 * @param {string|Object} [ops] Defaults to clearing contents and formats.
 * @returns {{success:boolean, canceled?:boolean, areas:Array, options:Object, cellCount:number, commentCount?:number, rowColStyleCount?:number, mergeCount?:number}}
 */
export function clearRange(range, ops) {
    const areas = _resolveAreas(this, range);
    const clearOptions = _normalizeClearOptions(ops);
    if (!areas.length || !_hasClearOption(clearOptions)) {
        return { success: false, areas, options: clearOptions, cellCount: 0 };
    }

    const beforeEvent = this.SN.Event.emit('beforeClearRange', { sheet: this, areas, options: clearOptions });
    if (beforeEvent.canceled) {
        return { success: false, canceled: true, areas, options: clearOptions, cellCount: 0 };
    }

    return this.SN._withOperation('clearRange', { sheet: this, areas, options: clearOptions }, () => {
        const changes = [];
        const comments = clearOptions.comments ? _snapshotComments(this, areas) : [];
        const mergeChange = _getMergeChanges(this, areas, clearOptions);
        const rowColStyleChanges = _collectRowColStyleChanges(this, areas, clearOptions);

        _forEachClearableCell(this, areas, clearOptions, (cell) => {
            if (_shouldMaterializeStyle(clearOptions)) cell.style;
            const before = _snapshotCell(cell);
            _clearCell(cell, clearOptions);
            const after = _snapshotCell(cell);
            if (!_sameSnapshot(before, after)) changes.push({ before, after });
        });

        if (clearOptions.comments) _clearComments(this, comments);
        if (mergeChange) _applyMerges(this, mergeChange.after);

        if (changes.length || comments.length || rowColStyleChanges.length || mergeChange) {
            this.SN.UndoRedo.add({
                undo: () => {
                    changes.forEach(change => _restoreCell(change.before));
                    _applyRowColStyleChanges(rowColStyleChanges, 'before');
                    if (clearOptions.comments) _restoreComments(this, comments);
                    if (mergeChange) _applyMerges(this, mergeChange.before);
                    _notifyAfterClear(this, changes, clearOptions);
                },
                redo: () => {
                    changes.forEach(change => _restoreCell(change.after));
                    _applyRowColStyleChanges(rowColStyleChanges, 'after');
                    if (clearOptions.comments) _clearComments(this, comments);
                    if (mergeChange) _applyMerges(this, mergeChange.after);
                    _notifyAfterClear(this, changes, clearOptions);
                }
            });
        }

        _notifyAfterClear(this, changes, clearOptions);
        const summary = {
            success: true,
            areas,
            options: clearOptions,
            cellCount: changes.length,
            commentCount: comments.length,
            rowColStyleCount: rowColStyleChanges.length,
            mergeCount: mergeChange?.removed?.length || 0
        };
        this.SN._recordChange({ type: 'range', action: 'clear', sheet: this, ...summary });
        this.SN.Event.emit('afterClearRange', { sheet: this, ...summary });
        return summary;
    });
}
