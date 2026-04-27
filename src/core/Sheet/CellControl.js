import { CELL_CONTROL_TYPES, isCheckboxControl, normalizeCellControl } from '../Cell/cellControlXml.js';

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

function _resolveAreas(sheet, range) {
    const fallback = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];
    const source = range ?? fallback;
    const list = Array.isArray(source) ? source : [source];
    const areas = [];

    list.forEach(item => {
        const parsed = typeof item === 'string' ? sheet.rangeStrToNum(item) : item;
        const normalized = _normalizeArea(parsed);
        if (normalized) areas.push(normalized);
    });
    return areas;
}

function _snapshotCell(cell) {
    return {
        cell,
        style: _clone(cell.style || {}),
        editVal: cell._editVal,
        calcVal: cell._calcVal,
        showVal: cell._showVal,
        fmtResult: cell._fmtResult,
        type: cell._type,
        richText: _clone(cell._richText)
    };
}

function _restoreCell(snapshot) {
    snapshot.cell._style = _clone(snapshot.style || {});
    snapshot.cell._editVal = snapshot.editVal;
    snapshot.cell._calcVal = snapshot.calcVal;
    snapshot.cell._showVal = snapshot.showVal;
    snapshot.cell._fmtResult = snapshot.fmtResult;
    snapshot.cell._type = snapshot.type;
    snapshot.cell._richText = _clone(snapshot.richText);
    _notifyCellChange(snapshot.cell);
}

function _applyBooleanValue(cell, value) {
    cell._editVal = !!value;
    cell._type = 'boolean';
    cell._calcVal = undefined;
    cell._showVal = undefined;
    cell._fmtResult = undefined;
    cell._richText = null;
}

function _clearCellValue(cell) {
    cell._editVal = '';
    cell._type = undefined;
    cell._calcVal = undefined;
    cell._showVal = undefined;
    cell._fmtResult = undefined;
    cell._richText = null;
}

function _getCheckboxDefaultValue(control) {
    if (control.default === 2) return '';
    return control.default === 1;
}

function _notifyCellChange(cell) {
    const sheet = cell.row.sheet;
    const SN = sheet.SN;
    if (SN.DependencyGraph) {
        SN.DependencyGraph.notifyChange(sheet.name, cell.row.rIndex, cell.cIndex);
    }
    if (sheet.Sparkline) sheet.Sparkline.clearAllCache();
    if (sheet.CF) sheet.CF.clearAllCache();
    if (sheet.Drawing) sheet.Drawing._clearChartRefCaches();
}

function _emitBeforeStyle(sheet, cell, oldStyle, newStyle) {
    if (!sheet.SN.Event.hasListeners('beforeCellStyleChange')) return false;
    const beforeEvent = sheet.SN.Event.emit('beforeCellStyleChange', {
        sheet,
        cell,
        row: cell.row.rIndex,
        col: cell.cIndex,
        oldStyle,
        newStyle
    });
    return beforeEvent.canceled;
}

function _emitAfterStyle(sheet, cell, oldStyle, newStyle) {
    if (!sheet.SN.Event.hasListeners('afterCellStyleChange')) return;
    sheet.SN.Event.emit('afterCellStyleChange', {
        sheet,
        cell,
        row: cell.row.rIndex,
        col: cell.cIndex,
        oldStyle,
        newStyle
    });
}

function _emitBeforeEdit(sheet, cell, newValue) {
    const oldValue = cell._editVal;
    const beforeEvent = sheet.SN.Event.emit('beforeCellEdit', {
        cell,
        sheet,
        row: cell.row.rIndex,
        col: cell.cIndex,
        oldValue,
        newValue
    });
    return beforeEvent.canceled;
}

function _emitAfterEdit(sheet, cell, oldValue, newValue) {
    sheet.SN.Event.emit('afterCellEdit', {
        cell,
        sheet,
        row: cell.row.rIndex,
        col: cell.cIndex,
        oldValue,
        newValue
    });
}

function _setControlOnCell(sheet, cell, control) {
    const oldStyle = _clone(cell.style || {});
    const newStyle = _clone(oldStyle || {});
    newStyle.control = _clone(control);

    let nextValue;
    let valueWillChange = false;
    if (!cell.isFormula && control.type === CELL_CONTROL_TYPES.CHECKBOX) {
        const current = cell.calcVal;
        nextValue = typeof current === 'boolean'
            ? current
            : _getCheckboxDefaultValue(control);
        valueWillChange = nextValue === ''
            ? cell._editVal !== '' || cell._type !== undefined
            : cell._editVal !== nextValue || cell._type !== 'boolean';
    }

    if (_emitBeforeStyle(sheet, cell, oldStyle, newStyle)) return null;
    if (valueWillChange && _emitBeforeEdit(sheet, cell, nextValue)) return null;

    const before = _snapshotCell(cell);
    cell._style = newStyle;
    if (valueWillChange) {
        if (nextValue === '') _clearCellValue(cell);
        else _applyBooleanValue(cell, nextValue);
        _notifyCellChange(cell);
    }
    const after = _snapshotCell(cell);

    _emitAfterStyle(sheet, cell, oldStyle, newStyle);
    if (valueWillChange) _emitAfterEdit(sheet, cell, before.editVal, nextValue);
    return { before, after };
}

function _clearControlOnCell(sheet, cell) {
    if (!isCheckboxControl(cell.control)) return null;

    const oldStyle = _clone(cell.style || {});
    const newStyle = _clone(oldStyle || {});
    delete newStyle.control;

    if (_emitBeforeStyle(sheet, cell, oldStyle, newStyle)) return null;

    const before = _snapshotCell(cell);
    cell._style = newStyle;
    const after = _snapshotCell(cell);

    _emitAfterStyle(sheet, cell, oldStyle, newStyle);
    return { before, after };
}

function _setCheckboxValueOnCell(sheet, cell, checked) {
    if (!isCheckboxControl(cell.control) || cell.isFormula) return null;
    const nextValue = !!checked;
    if (cell.calcVal === nextValue && cell._type === 'boolean') return null;
    if (_emitBeforeEdit(sheet, cell, nextValue)) return null;

    const before = _snapshotCell(cell);
    _applyBooleanValue(cell, nextValue);
    _notifyCellChange(cell);
    const after = _snapshotCell(cell);

    _emitAfterEdit(sheet, cell, before.editVal, nextValue);
    return { before, after };
}

function _commitCellControlChanges(sheet, operationName, areas, changes, record) {
    if (!changes.length) return { success: true, areas, cellCount: 0 };

    sheet.SN.UndoRedo.add({
        undo: () => changes.forEach(item => _restoreCell(item.before)),
        redo: () => changes.forEach(item => _restoreCell(item.after))
    });
    sheet.SN._recordChange(record);
    return { success: true, areas, cellCount: changes.length, operationName };
}

/** @param {string|Object|Array<string|Object>} range @param {Object} control */
export function setCellControl(range, control) {
    const areas = _resolveAreas(this, range);
    const normalized = normalizeCellControl(control);
    if (!areas.length || !normalized) return { success: false, canceled: true, cellCount: 0 };

    const beforeEvent = this.SN.Event.emit('beforeSetCellControl', {
        sheet: this,
        areas,
        control: _clone(normalized)
    });
    if (beforeEvent.canceled) return { success: false, canceled: true, cellCount: 0 };

    return this.SN._withOperation('setCellControl', { sheet: this, areas, control: _clone(normalized) }, () => {
        const changes = [];
        this.eachCells(areas, (r, c) => {
            const change = _setControlOnCell(this, this.getCell(r, c), normalized);
            if (change) changes.push(change);
        });

        const summary = _commitCellControlChanges(this, 'setCellControl', areas, changes, {
            type: 'cellControl',
            action: 'set',
            sheet: this,
            areas,
            control: _clone(normalized),
            cellCount: changes.length
        });
        this.SN.Event.emit('afterSetCellControl', { sheet: this, ...summary, control: _clone(normalized) });
        return summary;
    });
}

/** @param {string|Object|Array<string|Object>} range */
export function clearCellControl(range) {
    const areas = _resolveAreas(this, range);
    if (!areas.length) return { success: false, canceled: true, cellCount: 0 };

    const beforeEvent = this.SN.Event.emit('beforeClearCellControl', { sheet: this, areas });
    if (beforeEvent.canceled) return { success: false, canceled: true, cellCount: 0 };

    return this.SN._withOperation('clearCellControl', { sheet: this, areas }, () => {
        const changes = [];
        this.eachCells(areas, (r, c) => {
            const change = _clearControlOnCell(this, this.getCell(r, c));
            if (change) changes.push(change);
        }, { sparse: true });

        const summary = _commitCellControlChanges(this, 'clearCellControl', areas, changes, {
            type: 'cellControl',
            action: 'clear',
            sheet: this,
            areas,
            cellCount: changes.length
        });
        this.SN.Event.emit('afterClearCellControl', { sheet: this, ...summary });
        return summary;
    });
}

/** @param {number} row @param {number} col @returns {boolean} */
export function toggleCheckbox(row, col) {
    const cell = this.getCell(row, col);
    if (!isCheckboxControl(cell.control) || cell.isFormula) return false;
    const checked = cell.calcVal === true;
    const result = this.setCheckboxValue({ s: { r: row, c: col }, e: { r: row, c: col } }, !checked);
    return result?.cellCount > 0;
}

/** @param {string|Object|Array<string|Object>} range @param {boolean} checked */
export function setCheckboxValue(range, checked) {
    const areas = _resolveAreas(this, range);
    if (!areas.length) return { success: false, canceled: true, cellCount: 0 };

    return this.SN._withOperation('setCheckboxValue', { sheet: this, areas, checked: !!checked }, () => {
        const changes = [];
        this.eachCells(areas, (r, c) => {
            const change = _setCheckboxValueOnCell(this, this.getCell(r, c), checked);
            if (change) changes.push(change);
        }, { sparse: true });

        return _commitCellControlChanges(this, 'setCheckboxValue', areas, changes, {
            type: 'cellControl',
            action: 'setValue',
            sheet: this,
            areas,
            checked: !!checked,
            cellCount: changes.length
        });
    });
}

/** @param {string|Object|Array<string|Object>} range */
export function deleteCheckboxControlOrValue(range) {
    const areas = _resolveAreas(this, range);
    if (!areas.length) return { success: false, canceled: true, cellCount: 0 };

    let hasCheckbox = false;
    let hasChecked = false;
    this.eachCells(areas, (r, c) => {
        const cell = this.getCell(r, c);
        if (!isCheckboxControl(cell.control)) return;
        hasCheckbox = true;
        if (cell.calcVal === true) hasChecked = true;
    }, { sparse: true });

    if (!hasCheckbox) return { success: false, canceled: false, cellCount: 0 };
    return hasChecked ? this.setCheckboxValue(areas, false) : this.clearCellControl(areas);
}

/** @param {string|Object|Array<string|Object>} range */
export function hasCheckboxControl(range) {
    const areas = _resolveAreas(this, range);
    let found = false;
    this.eachCells(areas, (r, c) => {
        if (found) return;
        found = isCheckboxControl(this.getCell(r, c).control);
    }, { sparse: true });
    return found;
}
