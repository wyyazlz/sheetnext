/**
 * Table Table Manager
 * All tables in management form
 */
import TableItem from './TableItem.js';
import { generateTableName, toRef } from './TableUtils.js';

function cloneRange(range) {
    if (!range?.s || !range?.e) return null;
    return {
        s: { r: range.s.r, c: range.s.c },
        e: { r: range.e.r, c: range.e.c }
    };
}

function cloneColumns(columns = []) {
    return columns.map(col => ({ ...col }));
}

function isRangeEqual(a, b) {
    return !!a?.s && !!a?.e && !!b?.s && !!b?.e &&
        a.s.r === b.s.r && a.s.c === b.s.c &&
        a.e.r === b.e.r && a.e.c === b.e.c;
}

function isNonEmptyValue(value) {
    return value !== undefined && value !== null && value !== '' && !Number.isNaN(value);
}

export default class Table {
    /**
     * @param {Sheet} sheet
     */
    constructor(sheet) {
        /** @type {Sheet} */
        this.sheet = sheet;
        this._tables = new Map();
    }

    /** @type {number} */
    get size() {
        return this._tables.size;
    }

    /** @returns {TableItem[]} */
    getAll() {
        return Array.from(this._tables.values());
    }

    /**
     * Get table from ID
     * @param {string} id
     * @returns {TableItem|null}
     */
    get(id) {
        return this._tables.get(id) || null;
    }

    /**
     * Get table by name
     * @param {string} name
     * @returns {TableItem|null}
     */
    getByName(name) {
        for (const table of this._tables.values()) {
            if (table.name === name || table.displayName === name) {
                return table;
            }
        }
        return null;
    }

    /**
     * Get tables containing specified cells
     * @param {number} row
     * @param {number} col
     * @returns {TableItem|null}
     */
    getTableAt(row, col) {
        for (const table of this._tables.values()) {
            if (table.contains(row, col)) {
                return table;
            }
        }
        return null;
    }

    /**
     * Add/create new table
     * @param {Object} options
     * @returns {TableItem}
     */
    add(options = {}) {
        // 生成唯一名称
        if (!options.name) {
            options.name = generateTableName(this._tables);
        }
        if (!options.displayName) {
            options.displayName = options.name;
        }

        if (this.sheet.SN.Event.hasListeners('beforeTableAdd')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeTableAdd', {
                sheet: this.sheet,
                options
            });
            if (beforeEvent.canceled) return null;
        }
        const table = new TableItem(this.sheet, options);
        this._tables.set(table.id, table);

        // 判断是否是导入（仅 XML/文件导入路径显式标记）
        const isImport = options._fromImport === true;

        // 同步列名（仅非导入时，导入时列名已从 XML 解析）
        if (!isImport && table.showHeaderRow && table.range) {
            table.syncColumnNames();
        }

        if (!isImport && table.showTotalsRow && table.range) {
            table.applyTotalsRow({ initializeDefaults: true, force: true });
        }

        // 初始化自动筛选（导入时静默同步，避免 vi 未初始化触发渲染）
        if (table.autoFilterEnabled && table.range) {
            this.sheet.AutoFilter.registerTableScope(table, {
                keepFilters: false,
                silent: isImport,
                autoFilterXml: options._tableAutoFilterXml || null,
                restoreHiddenRowsFromSheet: isImport
            });
        }

        if (this.sheet.SN.Event.hasListeners('afterTableAdd')) {
            this.sheet.SN.Event.emit('afterTableAdd', { sheet: this.sheet, table, options });
        }
        return table;
    }

    /**
     * Create Table from Selection
     * @param {Object} area - { s: {r, c}, e: {r, c} }
     * @param {Object} options
     * @returns {TableItem}
     */
    createFromSelection(area, options = {}) {
        const targetArea = {
            s: {
                r: Math.min(area.s.r, area.e.r),
                c: Math.min(area.s.c, area.e.c)
            },
            e: {
                r: Math.max(area.s.r, area.e.r),
                c: Math.max(area.s.c, area.e.c)
            }
        };
        const isEmptyHeaderValue = value => value === undefined || value === null || value.toString().trim() === '';

        // 检查选区是否与现有表格重叠
        for (const table of this._tables.values()) {
            if (this._isOverlapping(targetArea, table.range)) {
                throw new Error('选区与现有表格重叠');
            }
        }

        // 创建列定义
        const colCount = targetArea.e.c - targetArea.s.c + 1;
        const columns = [];
        const hasHeader = options.showHeaderRow !== false;

        for (let i = 0; i < colCount; i++) {
            const col = targetArea.s.c + i;
            const defaultName = `列${i + 1}`;
            let name = defaultName;

            // 如果有表头，从单元格读取列名
            if (hasHeader) {
                const cell = this.sheet.getCell(targetArea.s.r, col);
                const headerValue = cell.calcVal ?? cell.v ?? cell.editVal;
                if (!isEmptyHeaderValue(headerValue)) {
                    name = headerValue.toString();
                } else {
                    // Keep header cells and table metadata consistent for xlsx table export.
                    cell.editVal = defaultName;
                }
            }

            columns.push({ id: i + 1, name });
        }

        return this.add({
            ...options,
            rangeRef: toRef(targetArea),
            columns,
            showHeaderRow: hasHeader
        });
    }

    /**
     * Remove Table
     * @param {string} id - Form ID
     * @returns {boolean}
     */
    remove(id) {
        const table = this._tables.get(id);
        if (table) {
            if (this.sheet.SN.Event.hasListeners('beforeTableRemove')) {
                const beforeEvent = this.sheet.SN.Event.emit('beforeTableRemove', { sheet: this.sheet, table, id });
                if (beforeEvent.canceled) return false;
            }
            // 清除筛选
            if (table.autoFilterEnabled) {
                this.sheet.AutoFilter.unregisterTableScope(table.id);
            }
            const removed = this._tables.delete(id);
            if (removed && this.sheet.SN.Event.hasListeners('afterTableRemove')) {
                this.sheet.SN.Event.emit('afterTableRemove', { sheet: this.sheet, table, id });
            }
            return removed;
        }
        return false;
    }

    /**
     * Remove table by name
     * @param {string} name
     * @returns {boolean}
     */
    removeByName(name) {
        const table = this.getByName(name);
        if (table) {
            return this.remove(table.id);
        }
        return false;
    }

    /**
     * Update name index (internal method, called by TableItem.rename)
     * @param {string} oldName
     * @param {string} newName
     * @param {TableItem} table
     * @private
     */
    _updateIndex(oldName, newName, table) {
        // 当前实现使用 id 作为 key，所以只需确保没有名称冲突
        // 如果将来改为名称索引，这里需要更新 Map
    }

    /**
     * Check if the two ranges overlap.
     * @private
     */
    _isOverlapping(range1, range2) {
        if (!range1 || !range2) return false;
        return !(range1.e.r < range2.s.r || range1.s.r > range2.e.r ||
                 range1.e.c < range2.s.c || range1.s.c > range2.e.c);
    }

    _normalizeArea(area) {
        if (!area) return null;
        const s = area.s || area;
        const e = area.e || area;
        if (!Number.isInteger(s.r) || !Number.isInteger(s.c) ||
            !Number.isInteger(e.r) || !Number.isInteger(e.c)) {
            return null;
        }
        return {
            s: { r: Math.min(s.r, e.r), c: Math.min(s.c, e.c) },
            e: { r: Math.max(s.r, e.r), c: Math.max(s.c, e.c) }
        };
    }

    _cellHasValue(row, col) {
        const cell = this.sheet.rows[row]?.cells?.[col];
        return isNonEmptyValue(cell?._editVal);
    }

    _getAutoExpandEndRow(table, area, options = {}) {
        const range = table.range;
        if (!range || table.showTotalsRow) return -1;

        const insertRow = range.e.r + 1;
        if (area.e.r < insertRow || area.s.r > insertRow) return -1;
        if (area.e.c < range.s.c || area.s.c > range.e.c) return -1;

        const isSingleCell = area.s.r === area.e.r && area.s.c === area.e.c;
        if (isSingleCell && Object.prototype.hasOwnProperty.call(options, 'value')) {
            return isNonEmptyValue(options.value) ? area.s.r : -1;
        }

        const startCol = Math.max(area.s.c, range.s.c);
        const endCol = Math.min(area.e.c, range.e.c);
        let endRow = -1;
        const hasEditedValue = typeof options.hasEditedValue === 'function'
            ? options.hasEditedValue
            : (r, c) => this._cellHasValue(r, c);

        for (let r = insertRow; r <= area.e.r; r++) {
            for (let c = startCol; c <= endCol; c++) {
                if (hasEditedValue(r, c)) {
                    endRow = r;
                    break;
                }
            }
        }

        return endRow;
    }

    _canResizeWithoutOverlap(targetTable, nextRange) {
        for (const table of this._tables.values()) {
            if (table === targetTable) continue;
            if (this._isOverlapping(nextRange, table.range)) return false;
        }
        return true;
    }

    _resetTableRowCache(table) {
        table._stripeRowIndexCache = null;
        table._stripeCacheStartRow = -1;
        table._stripeCacheEndRow = -1;
        table._stripeCacheVisibilityVersion = -1;
    }

    _captureLinkedPivotCacheStates(table) {
        const states = [];
        this.sheet.SN._pivotCaches?.forEach(cache => {
            if (!cache?._isLinkedToTable?.(table)) return;
            states.push({
                cache,
                state: cache._captureStructureState?.()
            });
        });
        return states;
    }

    _captureAutoExpandStateWithCaches(table, pivotCaches) {
        return {
            range: cloneRange(table.range),
            columns: cloneColumns(table.columns),
            pivotCaches: pivotCaches.map(item => ({
                cache: item.cache,
                state: item.cache?._captureStructureState?.()
            }))
        };
    }

    _applyAutoExpandState(table, state) {
        if (!table || !state) return;

        table._ref = cloneRange(state.range);
        table.columns = cloneColumns(state.columns);
        this._resetTableRowCache(table);

        if (table.showTotalsRow && table.range) {
            table.applyTotalsRow({ initializeDefaults: true, force: true });
        }

        if (table.autoFilterEnabled && table.range) {
            table._updateAutoFilter();
        } else {
            this.sheet.AutoFilter.unregisterTableScope(table.id, { silent: true });
        }

        state.pivotCaches?.forEach(item => {
            item.cache?._applyStructureState?.(item.state);
        });
    }

    _expandTableToRow(table, row, options = {}) {
        const oldRange = cloneRange(table.range);
        if (!oldRange || row <= oldRange.e.r) return false;

        const nextRange = cloneRange(oldRange);
        nextRange.e.r = row;
        if (!this._canResizeWithoutOverlap(table, nextRange)) return false;

        const linkedPivotCaches = this._captureLinkedPivotCacheStates(table);
        const beforeState = this._captureAutoExpandStateWithCaches(table, linkedPivotCaches);
        table.resize(nextRange);

        if (!isRangeEqual(table.range, nextRange)) return false;

        linkedPivotCaches.forEach(item => item.cache?._syncTableSource?.(table, { force: true }));
        const afterState = this._captureAutoExpandStateWithCaches(table, linkedPivotCaches);
        if (options.recordUndo !== false) {
            this.sheet.SN.UndoRedo.add({
                undo: () => this._applyAutoExpandState(table, beforeState),
                redo: () => this._applyAutoExpandState(table, afterState)
            });
        }

        this.sheet.SN._recordChange({
            type: 'tableAutoExpand',
            sheet: this.sheet,
            table,
            oldRange,
            newRange: cloneRange(table.range)
        });

        return true;
    }

    _autoExpandForEditArea(area, options = {}) {
        const normalizedArea = this._normalizeArea(area);
        if (!normalizedArea || this._tables.size === 0) return false;

        let changed = false;
        for (const table of this._tables.values()) {
            const endRow = this._getAutoExpandEndRow(table, normalizedArea, options);
            if (endRow < 0) continue;
            changed = this._expandTableToRow(table, endRow, options) || changed;
        }

        return changed;
    }

    _autoExpandForCellEdit(row, col, options = {}) {
        return this._autoExpandForEditArea({ s: { r: row, c: col }, e: { r: row, c: col } }, options);
    }

    /**
     * Walk through all tables
     * @param {Function} callback
     */
    forEach(callback) {
        this._tables.forEach(callback);
    }

    /**
     * Convert to JSON array
     */
    toJSON() {
        return this.getAll().map(table => table.toJSON());
    }

    /**
     * Restore from JSON
     * @param {Array} jsonArray
     */
    fromJSON(jsonArray) {
        // 清空现有表格
        this._tables.forEach(table => {
            if (table.autoFilterEnabled) {
                this.sheet.AutoFilter.unregisterTableScope(table.id, { silent: true });
            }
        });
        this._tables.clear();

        if (Array.isArray(jsonArray)) {
            jsonArray.forEach(json => {
                const table = TableItem.fromJSON(this.sheet, json);
                this._tables.set(table.id, table);
                if (table.autoFilterEnabled && table.range) {
                    this.sheet.AutoFilter.registerTableScope(table, { keepFilters: false, silent: true });
                }
            });
        }
    }
}
