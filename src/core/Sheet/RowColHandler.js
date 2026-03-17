/**
 * Line Operations Processing Module
 * Inserting and Deleting Handling Lines
 */

/**
 * Line Operations Processing Module
 * Inserting and Deleting Handling Lines
 */

function _cloneRange(range) {
    if (!range?.s || !range?.e) return null;
    return {
        s: { r: range.s.r, c: range.s.c },
        e: { r: range.e.r, c: range.e.c }
    };
}
function _cloneTableColumns(columns = []) {
    return columns.map(col => ({ ...col }));
}

function _buildNormalizedTableColumns(columns = [], count = columns.length) {
    const next = _cloneTableColumns(columns).slice(0, count);
    while (next.length < count) {
        next.push({
            id: next.length + 1,
            name: `列${next.length + 1}`
        });
    }
    return next.map((col, index) => ({
        ...col,
        id: index + 1,
        name: col?.name || `列${index + 1}`
    }));
}

function _resetTableRowCache(table) {
    table._stripeRowIndexCache = null;
    table._stripeCacheStartRow = -1;
    table._stripeCacheEndRow = -1;
    table._stripeCacheVisibilityVersion = -1;
}

function _applyTableSnapshots(sheet, snapshots, stage) {
    if (!sheet.Table || snapshots.length === 0) return;

    snapshots.forEach(snapshot => {
        const state = snapshot[stage];
        const table = snapshot.table;
        if (!state) return;

        if (!state.exists) {
            table._ref = null;
            sheet.Table._tables.delete(table.id);
            sheet.AutoFilter.unregisterTableScope(table.id, { silent: true });
            _resetTableRowCache(table);
            return;
        }

        table._ref = _cloneRange(state.range);
        table.columns = _cloneTableColumns(state.columns);
        table._showTotalsRow = state.showTotalsRow;
        sheet.Table._tables.set(table.id, table);
        _resetTableRowCache(table);

        if (table.showTotalsRow && table.range) {
            table.applyTotalsRow({ initializeDefaults: true, force: true });
        }

        if (table.autoFilterEnabled && table.range) {
            table._updateAutoFilter();
        } else {
            sheet.AutoFilter.unregisterTableScope(table.id, { silent: true });
        }
    });
}

function _buildRowDeleteTableSnapshots(sheet, r, number) {
    if (!sheet.Table || sheet.Table.size === 0) return [];

    const delEnd = r + number - 1;
    const snapshots = [];

    sheet.Table.forEach(table => {
        const range = table.range;
        if (!range || range.e.r < r) return;

        const before = {
            exists: true,
            range: _cloneRange(range),
            columns: _cloneTableColumns(table.columns),
            showTotalsRow: table.showTotalsRow
        };

        let exists = true;
        let nextRange = _cloneRange(range);
        let nextShowTotalsRow = table.showTotalsRow;

        if (range.s.r > delEnd) {
            nextRange.s.r -= number;
            nextRange.e.r -= number;
        } else {
            const nextStart = range.s.r < r ? range.s.r : (range.e.r > delEnd ? r : null);
            const nextEnd = range.e.r > delEnd ? range.e.r - number : (range.s.r < r ? r - 1 : null);

            if (nextStart === null || nextEnd === null || nextStart > nextEnd) {
                exists = false;
                nextRange = null;
            } else {
                nextRange.s.r = nextStart;
                nextRange.e.r = nextEnd;
            }
        }

        if (exists && nextRange) {
            const rowCount = nextRange.e.r - nextRange.s.r + 1;
            const minRows = (table.showHeaderRow ? 1 : 0) + (table.showTotalsRow ? 1 : 0);
            if (rowCount <= 0) {
                exists = false;
                nextRange = null;
            } else if (table.showTotalsRow && rowCount < minRows) {
                nextShowTotalsRow = false;
            }
        }

        snapshots.push({
            table,
            before,
            after: {
                exists,
                range: nextRange,
                columns: _cloneTableColumns(before.columns),
                showTotalsRow: nextShowTotalsRow
            }
        });
    });

    return snapshots;
}

function _buildColDeleteTableSnapshots(sheet, c, number) {
    if (!sheet.Table || sheet.Table.size === 0) return [];

    const delEnd = c + number - 1;
    const snapshots = [];

    sheet.Table.forEach(table => {
        const range = table.range;
        if (!range || range.e.c < c) return;

        const before = {
            exists: true,
            range: _cloneRange(range),
            columns: _cloneTableColumns(table.columns),
            showTotalsRow: table.showTotalsRow
        };

        let exists = true;
        let nextRange = _cloneRange(range);

        if (range.s.c > delEnd) {
            nextRange.s.c -= number;
            nextRange.e.c -= number;
        } else {
            const nextStart = range.s.c < c ? range.s.c : (range.e.c > delEnd ? c : null);
            const nextEnd = range.e.c > delEnd ? range.e.c - number : (range.s.c < c ? c - 1 : null);

            if (nextStart === null || nextEnd === null || nextStart > nextEnd) {
                exists = false;
                nextRange = null;
            } else {
                nextRange.s.c = nextStart;
                nextRange.e.c = nextEnd;
            }
        }

        let nextColumns = _cloneTableColumns(before.columns);
        if (exists && nextRange) {
            const overlapStart = Math.max(range.s.c, c);
            const overlapEnd = Math.min(range.e.c, delEnd);
            const removedCount = overlapStart <= overlapEnd ? (overlapEnd - overlapStart + 1) : 0;
            const removeIndex = Math.max(0, c - range.s.c);
            if (removedCount > 0) {
                nextColumns.splice(removeIndex, removedCount);
            }

            const colCount = nextRange.e.c - nextRange.s.c + 1;
            if (colCount <= 0) {
                exists = false;
                nextRange = null;
            } else {
                nextColumns = _buildNormalizedTableColumns(nextColumns, colCount);
            }
        }

        snapshots.push({
            table,
            before,
            after: {
                exists,
                range: nextRange,
                columns: exists ? nextColumns : [],
                showTotalsRow: table.showTotalsRow
            }
        });
    });

    return snapshots;
}

function _captureStructureStates(sheet) {
    return {
        pivotTable: sheet.PivotTable?._captureStructureState?.() ?? null,
        comment: sheet.Comment?._captureState?.() ?? null,
        crossSheet: sheet.SN.sheets.map(targetSheet => ({
            sheet: targetSheet,
            cf: targetSheet.CF?._captureState?.() ?? null,
            sparkline: targetSheet.Sparkline?._captureState?.() ?? null
        }))
    };
}

function _applyStructureStates(sheet, states) {
    if (!states) return;

    sheet.PivotTable?._applyStructureState?.(states.pivotTable);
    sheet.Comment?._applyState?.(states.comment);

    states.crossSheet?.forEach(entry => {
        entry.sheet?.CF?._applyState?.(entry.cf);
        entry.sheet?.Sparkline?._applyState?.(entry.sparkline);
    });
}

function _notifyStructureChanges(sheet, method, index, number) {
    sheet.PivotTable?.[method]?.(index, number);
    sheet.Comment?.[method]?.(index, number);
    sheet.SN.sheets.forEach(targetSheet => {
        targetSheet.CF?.[method]?.(index, number, sheet.name);
        targetSheet.Sparkline?.[method]?.(index, number, sheet.name);
    });
}

/**
 * Insert Row
 * @param {Number} r - Row Index to Insert Location
 * @param {Number} number - Number of rows inserted
 */
export function addRows(r, number = 1) {
    // 触发 beforeInsertRows 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeInsertRows', {
        sheet: this,
        row: r,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('insertRows', {
        sheet: this,
        row: r,
        count: number
    }, () => {
    // 保存受影响的合并单元格信息,用于撤销
    const cds = this.merges.filter(m => m.s.r < r && m.e.r >= r);
    const ams = this.merges.filter(m => m.s.r >= r);
    const oldMerges = [...cds, ...ams].map(m => ({ ...m })); // 深拷贝保存原始状态
    const oldStyledRows = new Set(this._styledRows); // 保存旧的样式行索引
    // 保存受影响的超级表信息（用于撤销）
    const affectedTables = [];
    const autoFilterBefore = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const structureBefore = _captureStructureStates(this);
    if (this.Table && this.Table.size > 0) {
        this.Table.forEach(table => {
            const range = table.range;
            if (range && r >= range.s.r && r <= range.e.r + 1) {
                // 插入位置在表格范围内或紧邻底部时，按 Excel 行为扩展表格高度
                affectedTables.push({
                    table,
                    oldRange: { s: { ...range.s }, e: { ...range.e } }
                });
            }
        });
    }

    const applyRowInsertToTables = () => {
        affectedTables.forEach(({ table, oldRange }) => {
            table._ref = {
                s: { ...oldRange.s },
                e: { ...oldRange.e, r: oldRange.e.r + number }
            };
            if (table.autoFilterEnabled) {
                table._updateAutoFilter();
            }
        });
    };
    const restoreTables = () => {
        affectedTables.forEach(({ table, oldRange }) => {
            table._ref = {
                s: { ...oldRange.s },
                e: { ...oldRange.e }
            };
            if (table.autoFilterEnabled) {
                table._updateAutoFilter();
            }
        });
    };

    this.rows.splice(r, 0, ...new Array(number));
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledRows 索引
    const newStyledRows = new Set();
    this._styledRows.forEach(idx => {
        if (idx >= r) newStyledRows.add(idx + number);
        else newStyledRows.add(idx);
    });
    this._styledRows = newStyledRows;

    // 更新受影响的合并单元格
    cds.forEach(m => {
        m.e.r += number;
        for (let i = r; i < r + number; i++) {
            for (let j = m.s.c; j <= m.e.c; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });

    ams.forEach(m => {
        m.s.r += number;
        m.e.r += number;
        for (let i = m.s.r; i <= m.e.r; i++) {
            for (let j = m.s.c; j <= m.e.c; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });
    applyRowInsertToTables();
    this.AutoFilter?._onRowsInserted?.(r, number);
    _notifyStructureChanges(this, '_onRowsInserted', r, number);
    const autoFilterAfter = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const structureAfter = _captureStructureStates(this);

    // 添加撤销重做操作
    const newStyledRowsAfter = new Set(this._styledRows); // 保存新的样式行索引
    this.SN.UndoRedo.add({
        undo: () => {
            // 删除插入的行
            this.rows.splice(r, number);
            // 恢复样式行索引
            this._styledRows = new Set(oldStyledRows);

            // 恢复合并单元格到原始状态
            [...cds, ...ams].forEach((m, i) => {
                Object.assign(m, oldMerges[i]);
                // 更新单元格的合并状态
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });
            restoreTables();
            this.AutoFilter?._applySheetScopeState?.(autoFilterBefore, { silent: true });
            _applyStructureStates(this, structureBefore);
        },
        redo: () => {
            // 重新插入行
            this.rows.splice(r, 0, ...new Array(number));
            // 恢复样式行索引
            this._styledRows = new Set(newStyledRowsAfter);

            // 更新合并单元格
            cds.forEach(m => {
                m.e.r += number;
                for (let i = r; i < r + number; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            ams.forEach(m => {
                m.s.r += number;
                m.e.r += number;
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });
            applyRowInsertToTables();
            this.AutoFilter?._applySheetScopeState?.(autoFilterAfter, { silent: true });
            _applyStructureStates(this, structureAfter);
        }
    });

    // 调整图纸位置（行插入）
    this.Drawing._onRowsInserted(r, number);

    // 触发 afterInsertRows 事件
    this.SN.Event.emit('afterInsertRows', {
        sheet: this,
        row: r,
        count: number
    });
    this.SN._recordChange({
        type: 'insertRows',
        sheet: this,
        row: r,
        count: number
    });
    });
}

/**
 * Insert Columns
 * @param {Number} c - Index of columns for inserting positions
 * @param {Number} number - Number of columns inserted
 */
export function addCols(c, number = 1) {
    // 触发 beforeInsertColumns 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeInsertColumns', {
        sheet: this,
        col: c,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('insertColumns', {
        sheet: this,
        col: c,
        count: number
    }, () => {
    // 保存受影响的合并单元格信息,用于撤销
    const cdsc = this.merges.filter(m => m.s.c < c && m.e.c >= c);
    const amc = this.merges.filter(m => m.s.c >= c);
    const oldMerges = [...cdsc, ...amc].map(m => ({ ...m })); // 深拷贝保存原始状态
    const oldStyledCols = new Set(this._styledCols); // 保存旧的样式列索引

    // 保存受影响的超级表信息（用于撤销）
    const affectedTables = [];
    if (this.Table && this.Table.size > 0) {
        this.Table.forEach(table => {
            const range = table.range;
            if (range && c >= range.s.c && c <= range.e.c + 1) {
                // 插入位置在表格范围内或紧邻右侧
                affectedTables.push({
                    table,
                    oldRange: { s: { ...range.s }, e: { ...range.e } },
                    oldColumns: table.columns.map(col => ({ ...col }))
                });
            }
        });
    }

    // 保存受影响的自动筛选信息（非超级表）
    const autoFilterBefore = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const structureBefore = _captureStructureStates(this);

    // 在每一行插入新列
    this.rows.forEach(row => {
        if (row) row.cells.splice(c, 0, ...new Array(number));
    });
    this.cols.splice(c, 0, ...new Array(number));
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledCols 索引
    const newStyledCols = new Set();
    this._styledCols.forEach(idx => {
        if (idx >= c) newStyledCols.add(idx + number);
        else newStyledCols.add(idx);
    });
    this._styledCols = newStyledCols;

    // 更新受影响的合并单元格的结束列
    cdsc.forEach(m => {
        m.e.c += number;
        // 给新单元格打上标记
        for (let i = m.s.r; i <= m.e.r; i++) {
            for (let j = c; j < c + number; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });

    amc.forEach(m => {
        m.s.c += number;
        m.e.c += number;
        for (let i = m.s.r; i <= m.e.r; i++) {
            for (let j = m.s.c; j <= m.e.c; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });

    // 更新受影响的超级表
    affectedTables.forEach(({ table, oldRange }) => {
        const range = table.range;
        const insertColIndex = c - range.s.c; // 相对于表格的插入位置

        // 扩展表格范围
        range.e.c += number;

        // 在插入位置添加新列定义
        for (let i = 0; i < number; i++) {
            table.columns.splice(insertColIndex, 0, {
                id: table.columns.length + i + 1,
                name: `列${insertColIndex + i + 1}`
            });
        }

        // 更新自动筛选范围
        if (table.autoFilterEnabled) {
            table._updateAutoFilter();
        }
    });

    // 更新普通自动筛选范围（非超级表）
    this.AutoFilter?._onColsInserted?.(c, number);
    _notifyStructureChanges(this, '_onColsInserted', c, number);
    const autoFilterAfter = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const structureAfter = _captureStructureStates(this);

    // 添加撤销重做操作
    const newStyledColsAfter = new Set(this._styledCols); // 保存新的样式列索引
    this.SN.UndoRedo.add({
        undo: () => {
            // 删除插入的列
            this.rows.forEach(row => {
                if (row) row.cells.splice(c, number);
            });
            this.cols.splice(c, number);
            // 恢复样式列索引
            this._styledCols = new Set(oldStyledCols);

            // 恢复合并单元格到原始状态
            [...cdsc, ...amc].forEach((m, i) => {
                Object.assign(m, oldMerges[i]);
                // 更新单元格的合并状态
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            // 恢复超级表到原始状态
            affectedTables.forEach(({ table, oldRange, oldColumns }) => {
                table._ref = oldRange;
                table.columns = oldColumns;
                if (table.autoFilterEnabled) {
                    table._updateAutoFilter();
                }
            });

            // 恢复自动筛选范围
            this.AutoFilter?._applySheetScopeState?.(autoFilterBefore, { silent: true });
            _applyStructureStates(this, structureBefore);
        },
        redo: () => {
            // 重新插入列
            this.rows.forEach(row => {
                if (row) row.cells.splice(c, 0, ...new Array(number));
            });
            this.cols.splice(c, 0, ...new Array(number));
            // 恢复样式列索引
            this._styledCols = new Set(newStyledColsAfter);

            // 更新合并单元格
            cdsc.forEach(m => {
                m.e.c += number;
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = c; j < c + number; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            amc.forEach(m => {
                m.s.c += number;
                m.e.c += number;
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            // 重新更新超级表
            affectedTables.forEach(({ table, oldRange }) => {
                const range = table.range;
                const insertColIndex = c - oldRange.s.c;
                range.e.c += number;
                for (let i = 0; i < number; i++) {
                    table.columns.splice(insertColIndex, 0, {
                        id: table.columns.length + i + 1,
                        name: `列${insertColIndex + i + 1}`
                    });
                }
                if (table.autoFilterEnabled) {
                    table._updateAutoFilter();
                }
            });

            // 重新更新自动筛选
            this.AutoFilter?._applySheetScopeState?.(autoFilterAfter, { silent: true });
            _applyStructureStates(this, structureAfter);
        }
    });

    // 调整图纸位置（列插入）
    this.Drawing._onColsInserted(c, number);

    // 触发 afterInsertColumns 事件
    this.SN.Event.emit('afterInsertColumns', {
        sheet: this,
        col: c,
        count: number
    });
    this.SN._recordChange({
        type: 'insertColumns',
        sheet: this,
        col: c,
        count: number
    });
    });
}

/**
 * Delete Row
 * @param {Number} r - Remove Row Index at Start Location
 * @param {Number} number - Number of rows deleted
 */
export function delRows(r, number = 1) {
    // 触发 beforeDeleteRows 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeDeleteRows', {
        sheet: this,
        row: r,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('deleteRows', {
        sheet: this,
        row: r,
        count: number
    }, () => {
    // 存储将要删除的行对象
    const deletedRows = this.rows.slice(r, r + number);
    // 存储删除前的合并单元格信息
    const mergesBefore = this.merges.slice();
    const oldStyledRows = new Set(this._styledRows); // 保存旧的样式行索引
    const tableSnapshots = _buildRowDeleteTableSnapshots(this, r, number);
    const autoFilterBefore = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const structureBefore = _captureStructureStates(this);

    // 检测并更新涉及合并单元格
    this.merges = this.merges.filter(m => {
        // 如果删除区域完全在合并单元格之前，无需修改
        if (m.e.r < r) return true;
        // 如果删除区域完全在合并单元格之后，只需更新索引
        if (m.s.r >= r + number) {
            this.eachCells(m, (rowIndex, c) => {
                this.getCell(rowIndex, c).master = { r: m.s.r - number, c: m.s.c }
            })
            m.s.r -= number;
            m.e.r -= number;
            return true;
        }
        // 判断是否需要调整合并单元格的起始或结束行
        if (m.s.r == r) { // 如果删除的是起始行
            if (m.e.r - m.s.r < number) {
                // 如果删除的行数大于等于合并单元格的总行数，该合并单元格应被移除
                return false;
            } else {
                // 更新剩下的单元格的master行
                m.e.r -= number;
                return true;
            }
        } else { // 删除的是合并单元格中间或尾部的行
            m.e.r -= number;
            return true;
        }
    });

    // 删除指定的行
    this.rows.splice(r, number);
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledRows 索引
    const newStyledRows = new Set();
    this._styledRows.forEach(idx => {
        if (idx >= r && idx < r + number) return; // 删除的行，跳过
        if (idx >= r + number) newStyledRows.add(idx - number);
        else newStyledRows.add(idx);
    });
    this._styledRows = newStyledRows;

    // 如果有合并区域剩一个单元格，清除所有
    this.merges = this.merges.filter(m => {
        if (m.s.r == m.e.r && m.s.c == m.e.c) {
            const cell = this.getCell(m.s.r, m.s.c)
            cell.isMerged = false
            cell.master = null
            return false
        } else {
            return true
        }
    })

    // 存储删除后的合并单元格信息
    _applyTableSnapshots(this, tableSnapshots, 'after');
    this.AutoFilter?._onRowsDeleted?.(r, number);
    _notifyStructureChanges(this, '_onRowsDeleted', r, number);
    const mergesAfter = this.merges.slice();
    const autoFilterAfter = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const newStyledRowsAfter = new Set(this._styledRows);
    const structureAfter = _captureStructureStates(this);

    // 添加撤销重做功能
    this.SN.UndoRedo.add({
        undo: () => {
            // 撤销操作
            // 将删除的行插回原位置
            this.rows.splice(r, 0, ...deletedRows);
            // 恢复样式行索引
            this._styledRows = new Set(oldStyledRows);
            // 恢复删除前的合并单元格信息
            this.merges = mergesBefore.slice();
            _applyTableSnapshots(this, tableSnapshots, 'before');
            this.AutoFilter?._applySheetScopeState?.(autoFilterBefore, { silent: true });
            _applyStructureStates(this, structureBefore);
        },
        redo: () => {
            // 重做操作
            // 再次删除指定的行
            this.rows.splice(r, number);
            // 恢复样式行索引
            this._styledRows = new Set(newStyledRowsAfter);
            // 恢复删除后的合并单元格信息
            this.merges = mergesAfter.slice();
            _applyTableSnapshots(this, tableSnapshots, 'after');
            this.AutoFilter?._applySheetScopeState?.(autoFilterAfter, { silent: true });
            _applyStructureStates(this, structureAfter);
        }
    });

    // 调整图纸位置（行删除）
    this.Drawing._onRowsDeleted(r, number);

    // 触发 afterDeleteRows 事件
    this.SN.Event.emit('afterDeleteRows', {
        sheet: this,
        row: r,
        count: number
    });
    this.SN._recordChange({
        type: 'deleteRows',
        sheet: this,
        row: r,
        count: number
    });
    });
}

/**
 * Delete Column
 * @param {Number} c - Delete column index for starting position
 * @param {Number} number - NUMBER OF ROWS REMOVED
 */
export function delCols(c, number = 1) {
    // 触发 beforeDeleteColumns 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeDeleteColumns', {
        sheet: this,
        col: c,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('deleteColumns', {
        sheet: this,
        col: c,
        count: number
    }, () => {
    // 存储删除前的状态
    const deletedCols = this.cols.slice(c, c + number);
    const deletedCells = this.rows.map(row => row ? row.cells.slice(c, c + number) : null);
    const mergesBefore = this.merges.slice();
    const oldStyledCols = new Set(this._styledCols); // 保存旧的样式列索引
    const tableSnapshots = _buildColDeleteTableSnapshots(this, c, number);
    const autoFilterBefore = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const structureBefore = _captureStructureStates(this);

    // 检测并更新涉及合并单元格
    this.merges = this.merges.filter(m => {
        // 如果删除区域完全在合并单元格之前，无需修改
        if (m.e.c < c) return true;
        // 如果删除区域完全在合并单元格之后，只需更新索引
        if (m.s.c >= c + number) {
            this.eachCells(m, (r, c) => {
                this.getCell(r, c).master = { r: m.s.r, c: m.s.c - number }
            })
            m.s.c -= number;
            m.e.c -= number;
            return true;
        }
        if (m.s.c == c) { // 删的如果是master
            if (m.e.c - m.s.c < number) {
                return false
            } else { // 更新剩下的单元格的master
                m.e.c = m.e.c - number
                return true
            }
        } else { // 删非master
            m.e.c = m.e.c - number
            return true
        }
    });

    // 删除指定的列
    this.rows.forEach(row => {
        if (row) row.cells.splice(c, number);
    });
    this.cols.splice(c, number);
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledCols 索引
    const newStyledCols = new Set();
    this._styledCols.forEach(idx => {
        if (idx >= c && idx < c + number) return; // 删除的列，跳过
        if (idx >= c + number) newStyledCols.add(idx - number);
        else newStyledCols.add(idx);
    });
    this._styledCols = newStyledCols;

    // 如果剩一个单元格，清除所有
    this.merges = this.merges.filter(m => {
        if (m.s.r == m.e.r && m.s.c == m.e.c) {
            const cell = this.getCell(m.s.r, m.s.c)
            cell.isMerged = false
            cell.master = null
            return false
        } else {
            return true
        }
    })

    // 存储删除后的状态
    _applyTableSnapshots(this, tableSnapshots, 'after');
    this.AutoFilter?._onColsDeleted?.(c, number);
    _notifyStructureChanges(this, '_onColsDeleted', c, number);
    const mergesAfter = this.merges.slice();
    const autoFilterAfter = this.AutoFilter?._captureSheetScopeState?.() ?? null;
    const newStyledColsAfter = new Set(this._styledCols);
    const structureAfter = _captureStructureStates(this);

    // 添加撤销重做功能
    this.SN.UndoRedo.add({
        undo: () => {
            // 恢复列
            this.cols.splice(c, 0, ...deletedCols);
            // 恢复单元格数据
            this.rows.forEach((row, rowIndex) => {
                if (row && deletedCells[rowIndex]) {
                    row.cells.splice(c, 0, ...deletedCells[rowIndex]);
                }
            });
            // 恢复样式列索引
            this._styledCols = new Set(oldStyledCols);
            // 恢复合并单元格信息
            this.merges = mergesBefore.slice();
            _applyTableSnapshots(this, tableSnapshots, 'before');
            this.AutoFilter?._applySheetScopeState?.(autoFilterBefore, { silent: true });
            _applyStructureStates(this, structureBefore);
        },
        redo: () => {
            // 重新删除列
            this.rows.forEach(row => {
                if (row) row.cells.splice(c, number);
            });
            this.cols.splice(c, number);
            // 恢复样式列索引
            this._styledCols = new Set(newStyledColsAfter);
            // 恢复删除后的合并单元格信息
            this.merges = mergesAfter.slice();
            _applyTableSnapshots(this, tableSnapshots, 'after');
            this.AutoFilter?._applySheetScopeState?.(autoFilterAfter, { silent: true });
            _applyStructureStates(this, structureAfter);
        }
    });

    // 调整图纸位置（列删除）
    this.Drawing._onColsDeleted(c, number);

    // 触发 afterDeleteColumns 事件
    this.SN.Event.emit('afterDeleteColumns', {
        sheet: this,
        col: c,
        count: number
    });
    this.SN._recordChange({
        type: 'deleteColumns',
        sheet: this,
        col: c,
        count: number
    });
    });
}
