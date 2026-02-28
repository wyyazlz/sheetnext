import { TABLE_STYLES, resolveTableStyle } from './TableStyle.js';
import { parseRef, toRef, toCellRef } from './TableUtils.js';

const TOTALS_FUNCTION_TO_SUBTOTAL_CODE = {
    sum: 109,
    average: 101,
    count: 103,
    countNums: 102,
    max: 104,
    min: 105,
    stdDev: 107,
    var: 110
};

const TOTALS_FUNCTION_ALIASES = {
    sum: 'sum',
    average: 'average',
    count: 'count',
    countnums: 'countNums',
    max: 'max',
    min: 'min',
    stddev: 'stdDev',
    var: 'var'
};

/** @class */
export default class TableItem {
    /** @param {Sheet} sheet @param {Object} options */
    constructor(sheet, options = {}) {
        this.sheet = sheet;
        /** @type {string} */
        this.id = options.id || this._generateId();
        this._name = options.name || `表${this.id}`;
        this._displayName = options.displayName || this._name;

        const rangeRef = options.rangeRef ?? options.ref ?? null;
        this._ref = rangeRef ? parseRef(rangeRef) : null;

        /** @type {Array<Object>} */
        this.columns = options.columns || [];

        this._styleId = options.styleId || 'TableStyleMedium2';
        /** @type {boolean} */
        this.showHeaderRow = options.showHeaderRow ?? true;
        this._showTotalsRow = options.showTotalsRow ?? false;
        /** @type {boolean} */
        this.showFirstColumn = options.showFirstColumn ?? false;
        /** @type {boolean} */
        this.showLastColumn = options.showLastColumn ?? false;
        /** @type {boolean} */
        this.showRowStripes = options.showRowStripes ?? true;
        /** @type {boolean} */
        this.showColumnStripes = options.showColumnStripes ?? false;

        this._autoFilterEnabled = options.autoFilterEnabled ?? true;
        this._fromImport = options._fromImport === true;

        this._stripeRowIndexCache = null;
        this._stripeCacheStartRow = -1;
        this._stripeCacheEndRow = -1;
        this._stripeCacheVisibilityVersion = -1;

        /** @type {number|undefined} */
        this.headerRowDxfId = options.headerRowDxfId;
        /** @type {number|undefined} */
        this.dataDxfId = options.dataDxfId;
        /** @type {number|undefined} */
        this.headerRowBorderDxfId = options.headerRowBorderDxfId;
        /** @type {number|undefined} */
        this.tableBorderDxfId = options.tableBorderDxfId;
        /** @type {number|undefined} */
        this.totalsRowBorderDxfId = options.totalsRowBorderDxfId;

        // Theme-aware style cache
        this._resolvedStyleCache = null;
        this._resolvedStyleCacheStyleId = null;
        this._resolvedStyleCacheThemeRef = null;

        // dxf style cache
        this._dxfStyleCache = new Map();
    }

    _generateId() {
        return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
    }

    // ===================== Getters & Setters =====================

    /** @type {string} */
    get name() {
        return this._name;
    }

    set name(value) {
        if (!value || typeof value !== 'string') {
            throw new Error('请输入有效的表名');
        }

        const trimmedName = value.trim();
        if (!trimmedName) {
            throw new Error('表名不能为空');
        }

        const oldName = this._name;
        if (oldName === trimmedName) return;

        // 检查名称是否已被其他表格使用
        const existingTable = this.sheet.Table.getByName(trimmedName);
        if (existingTable && existingTable !== this) {
            throw new Error(`表名 "${trimmedName}" 已被使用`);
        }

        if (this.sheet.SN.Event.hasListeners('beforeTableRename')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeTableRename', {
                sheet: this.sheet,
                table: this,
                oldName,
                newName: trimmedName
            });
            if (beforeEvent.canceled) return;
        }
        this._name = trimmedName;
        this._displayName = trimmedName;

        this.sheet.Table._updateIndex(oldName, trimmedName, this);

        if (this.sheet.SN.Event.hasListeners('afterTableRename')) {
            this.sheet.SN.Event.emit('afterTableRename', {
                sheet: this.sheet,
                table: this,
                oldName,
                newName: trimmedName
            });
        }
    }

    /** @type {string} */
    get displayName() {
        return this._displayName;
    }

    set displayName(value) {
        this._displayName = value;
    }

    /** @type {string|null} */
    get styleId() {
        return this._styleId;
    }

    set styleId(value) {
        const oldStyle = this._styleId ?? null;
        let nextStyle = null;
        if (value === null) {
            nextStyle = null;
        } else if (TABLE_STYLES[value]) {
            nextStyle = value;
        } else {
            return;
        }
        if (oldStyle === nextStyle) return;
        if (this.sheet.SN.Event.hasListeners('beforeTableStyleChange')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeTableStyleChange', {
                sheet: this.sheet,
                table: this,
                oldStyle,
                newStyle: nextStyle
            });
            if (beforeEvent.canceled) return;
        }
        this._styleId = nextStyle;
        this._clearStyleDxfOverrides();
        if (this.sheet.SN.Event.hasListeners('afterTableStyleChange')) {
            this.sheet.SN.Event.emit('afterTableStyleChange', {
                sheet: this.sheet,
                table: this,
                oldStyle,
                newStyle: nextStyle
            });
        }
    }

    /** @type {boolean} */
    get showTotalsRow() {
        return this._showTotalsRow;
    }

    set showTotalsRow(value) {
        if (!this._ref) return;

        const next = !!value;
        if (this._showTotalsRow === next) return;

        const oldValue = this._showTotalsRow;
        if (this.sheet.SN.Event.hasListeners('beforeTableTotalsRowToggle')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeTableTotalsRowToggle', {
                sheet: this.sheet,
                table: this,
                oldValue,
                newValue: next
            });
            if (beforeEvent.canceled) return;
        }

        if (next) {
            this._ref.e.r += 1;
        } else {
            this._ref.e.r -= 1;
        }

        this._showTotalsRow = next;

        if (next) {
            this.applyTotalsRow({ initializeDefaults: true, force: true });
        }

        this._updateAutoFilter();
        if (this.sheet.SN.Event.hasListeners('afterTableTotalsRowToggle')) {
            this.sheet.SN.Event.emit('afterTableTotalsRowToggle', {
                sheet: this.sheet,
                table: this,
                oldValue,
                newValue: next
            });
        }
    }

    /** @type {boolean} */
    get autoFilterEnabled() {
        return this._autoFilterEnabled;
    }

    set autoFilterEnabled(value) {
        const next = !!value;
        const oldValue = this._autoFilterEnabled;
        if (oldValue === next) return;
        if (this.sheet.SN.Event.hasListeners('beforeTableAutoFilterToggle')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeTableAutoFilterToggle', {
                sheet: this.sheet,
                table: this,
                oldValue,
                newValue: next
            });
            if (beforeEvent.canceled) return;
        }
        this._autoFilterEnabled = next;
        if (next) {
            this._updateAutoFilter();
        } else {
            this.sheet.AutoFilter.unregisterTableScope(this.id);
        }
        if (this.sheet.SN.Event.hasListeners('afterTableAutoFilterToggle')) {
            this.sheet.SN.Event.emit('afterTableAutoFilterToggle', {
                sheet: this.sheet,
                table: this,
                oldValue,
                newValue: next
            });
        }
    }

    /** @type {string|null} */
    get ref() {
        return this._ref ? toRef(this._ref) : null;
    }

    set ref(value) {
        this._ref = typeof value === 'string' ? parseRef(value) : value;
    }

    /** @type {{s: {r: number, c: number}, e: {r: number, c: number}}|null} */
    get range() {
        return this._ref;
    }

    /** @type {number} */
    get headerRowIndex() {
        return this.showHeaderRow && this._ref ? this._ref.s.r : -1;
    }

    /** @type {number} */
    get dataStartRow() {
        return this._ref ? (this.showHeaderRow ? this._ref.s.r + 1 : this._ref.s.r) : -1;
    }

    /** @type {number} */
    get dataEndRow() {
        return this._ref ? (this.showTotalsRow ? this._ref.e.r - 1 : this._ref.e.r) : -1;
    }

    /** @type {number} */
    get totalsRowIndex() {
        return this.showTotalsRow && this._ref ? this._ref.e.r : -1;
    }

    /** @type {Object|null} */
    get style() {
        if (this._styleId === null) return null;

        const themeRef = this.sheet?.SN?.Xml?.clrScheme || this.sheet?.SN?.Xml?.clrSchemeTp || null;
        if (this._resolvedStyleCache &&
            this._resolvedStyleCacheStyleId === this._styleId &&
            this._resolvedStyleCacheThemeRef === themeRef) {
            return this._resolvedStyleCache;
        }

        const resolved = resolveTableStyle(this._styleId, this.sheet?.SN) || TABLE_STYLES.TableStyleMedium2;
        this._resolvedStyleCache = resolved;
        this._resolvedStyleCacheStyleId = this._styleId;
        this._resolvedStyleCacheThemeRef = themeRef;
        return resolved;
    }

    // ===================== Core Methods =====================

    /** @param {number} row @param {number} col @returns {boolean} */
    contains(row, col) {
        if (!this._ref) return false;
        const { s, e } = this._ref;
        return row >= s.r && row <= e.r && col >= s.c && col <= e.c;
    }

    /** @param {number} row @param {number} col @returns {boolean} */
    isHeaderCell(row, col) {
        return this.showHeaderRow && this.contains(row, col) && row === this._ref.s.r;
    }

    /** @param {number} row @param {number} col @returns {boolean} */
    isTotalsCell(row, col) {
        return this.showTotalsRow && this.contains(row, col) && row === this._ref.e.r;
    }

    /** @param {number} row @param {number} col @returns {boolean} */
    isDataCell(row, col) {
        if (!this.contains(row, col)) return false;
        if (this.isHeaderCell(row, col)) return false;
        if (this.isTotalsCell(row, col)) return false;
        return true;
    }

    /** @param {number} row @returns {number} visible row index in data area, -1 if hidden or outside */
    getDataRowIndex(row) {
        if (!this._ref) return -1;
        const startRow = this.dataStartRow;
        const endRow = this.dataEndRow;
        if (row < startRow || row > endRow) return -1;

        const visibilityVersion = this.sheet._rowVisibilityVersion || 0;
        const cacheMiss = !this._stripeRowIndexCache ||
            this._stripeCacheStartRow !== startRow ||
            this._stripeCacheEndRow !== endRow ||
            this._stripeCacheVisibilityVersion !== visibilityVersion;

        if (cacheMiss) {
            const cacheLength = endRow - startRow + 1;
            const cache = new Int32Array(cacheLength);
            cache.fill(-1);

            let visibleIndex = 0;
            for (let r = startRow; r <= endRow; r++) {
                if (!this.sheet.getRow(r).hidden) {
                    cache[r - startRow] = visibleIndex;
                    visibleIndex += 1;
                }
            }

            this._stripeRowIndexCache = cache;
            this._stripeCacheStartRow = startRow;
            this._stripeCacheEndRow = endRow;
            this._stripeCacheVisibilityVersion = visibilityVersion;
        }

        return this._stripeRowIndexCache[row - startRow] ?? -1;
    }

    /** @param {number} col @returns {number} */
    getColumnIndex(col) {
        if (!this._ref) return -1;
        if (col < this._ref.s.c || col > this._ref.e.c) return -1;
        return col - this._ref.s.c;
    }

    /** @param {number} colIndex @returns {Object|null} */
    getColumn(colIndex) {
        return this.columns[colIndex] || null;
    }

    _normalizeDxfId(dxfId) {
        if (dxfId === null || dxfId === undefined || dxfId === '') return null;
        const parsed = Number.parseInt(dxfId, 10);
        if (!Number.isInteger(parsed) || parsed < 0) return null;
        return parsed;
    }

    _getDxfStyle(dxfId) {
        const normalized = this._normalizeDxfId(dxfId);
        if (normalized === null) return null;

        if (this._dxfStyleCache.has(normalized)) {
            return this._dxfStyleCache.get(normalized);
        }

        const dxfStyle = this.sheet?.SN?.Xml?._xGetDxfStyle?.(normalized) || null;
        this._dxfStyleCache.set(normalized, dxfStyle);
        return dxfStyle;
    }

    _getHeaderDxfStyle() {
        return this._getDxfStyle(this.headerRowDxfId);
    }

    _getDataDxfStyle(colIndex) {
        const column = this.getColumn(colIndex);
        const dxfId = column?.dataDxfId ?? this.dataDxfId;
        return this._getDxfStyle(dxfId);
    }

    _clearStyleDxfOverrides() {
        let changed = false;
        const clearDxfProp = (prop) => {
            if (this[prop] === undefined || this[prop] === null || this[prop] === '') return;
            this[prop] = undefined;
            changed = true;
        };

        clearDxfProp('headerRowDxfId');
        clearDxfProp('dataDxfId');
        clearDxfProp('headerRowBorderDxfId');
        clearDxfProp('tableBorderDxfId');
        clearDxfProp('totalsRowBorderDxfId');

        for (const column of this.columns) {
            if (!column) continue;
            if (column.dataDxfId === undefined || column.dataDxfId === null || column.dataDxfId === '') continue;
            delete column.dataDxfId;
            changed = true;
        }

        if (changed) this._dxfStyleCache.clear();
    }

    /** @param {string} name @returns {number} */
    getColumnIndexByName(name) {
        return this.columns.findIndex(col => col.name === name);
    }

    /** @param {string|Object} newRef */
    resize(newRef) {
        const oldRef = this._ref;
        const nextRef = typeof newRef === 'string' ? parseRef(newRef) : newRef;
        if (this.sheet.SN.Event.hasListeners('beforeTableResize')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeTableResize', {
                sheet: this.sheet,
                table: this,
                oldRef,
                newRef: nextRef
            });
            if (beforeEvent.canceled) return;
        }
        this._ref = nextRef;

        // 更新列定义数量
        const colCount = this._ref.e.c - this._ref.s.c + 1;
        while (this.columns.length < colCount) {
            this.columns.push({
                id: this.columns.length + 1,
                name: `列${this.columns.length + 1}`
            });
        }
        if (this.columns.length > colCount) {
            this.columns = this.columns.slice(0, colCount);
        }

        // 更新自动筛选范围
        if (this.showTotalsRow) {
            this.applyTotalsRow({ initializeDefaults: true });
        }

        if (this.autoFilterEnabled) {
            this._updateAutoFilter();
        }
        if (this.sheet.SN.Event.hasListeners('afterTableResize')) {
            this.sheet.SN.Event.emit('afterTableResize', {
                sheet: this.sheet,
                table: this,
                oldRef,
                newRef: this._ref
            });
        }
    }

    /** Sync column names from header cell values. */
    syncColumnNames() {
        if (!this.showHeaderRow || !this._ref) return;

        const headerRow = this._ref.s.r;
        const colCount = this._ref.e.c - this._ref.s.c + 1;

        // 确保 columns 数组长度足够
        while (this.columns.length < colCount) {
            this.columns.push({
                id: this.columns.length + 1,
                name: `列${this.columns.length + 1}`
            });
        }

        for (let c = this._ref.s.c; c <= this._ref.e.c; c++) {
            const cell = this.sheet.getCell(headerRow, c);
            const colIndex = c - this._ref.s.c;
            // 使用 calcVal 获取实际显示值（而非原始值 v，因为字符串可能是 sst 索引）
            const cellValue = cell.calcVal ?? cell.v;
            this.columns[colIndex].name = cellValue?.toString() || `列${colIndex + 1}`;
        }
    }

    _updateAutoFilter() {
        if (!this.autoFilterEnabled || !this._ref) return;
        this.sheet.AutoFilter.registerTableScope(this, { keepFilters: true });
    }

    _normalizeTotalsFunction(funcType) {
        if (funcType == null) return null;
        const raw = String(funcType).trim();
        if (!raw) return null;
        const lowered = raw.toLowerCase();
        if (lowered === 'none' || lowered === 'null') return null;
        return TOTALS_FUNCTION_ALIASES[lowered] || null;
    }

    _getOrCreateColumn(colIndex) {
        while (this.columns.length <= colIndex) {
            this.columns.push({
                id: this.columns.length + 1,
                name: `\u5217${this.columns.length + 1}`
            });
        }

        const column = this.columns[colIndex];
        if (!column.id) column.id = colIndex + 1;
        if (!column.name) column.name = `\u5217${colIndex + 1}`;
        return column;
    }

    _isNumericTotalsCandidate(colIndex) {
        if (!this._ref || colIndex < 0) return false;
        if (this.dataEndRow < this.dataStartRow) return false;

        const absCol = this._ref.s.c + colIndex;
        let hasNumeric = false;

        for (let r = this.dataStartRow; r <= this.dataEndRow; r++) {
            const row = this.sheet.rows[r];
            const cell = row?.cells?.[absCol];
            if (!cell) continue;

            const cellType = cell.type;
            if (cellType === 'date' || cellType === 'dateTime' || cellType === 'time') continue;

            const value = cell.calcVal;
            if (typeof value === 'number' && Number.isFinite(value)) {
                hasNumeric = true;
            }
        }

        return hasNumeric;
    }

    _buildTotalsFormula(colIndex, funcType) {
        const normalized = this._normalizeTotalsFunction(funcType);
        if (!normalized) return '';

        const subtotalCode = TOTALS_FUNCTION_TO_SUBTOTAL_CODE[normalized];
        if (!subtotalCode) return '';

        if (this.dataEndRow < this.dataStartRow) {
            return '=0';
        }

        const absCol = this._ref.s.c + colIndex;
        const startRef = toCellRef(this.dataStartRow, absCol);
        const endRef = toCellRef(this.dataEndRow, absCol);
        return `=SUBTOTAL(${subtotalCode},${startRef}:${endRef})`;
    }

    _ensureTotalsDefaults() {
        if (!this._ref) return;

        const colCount = this._ref.e.c - this._ref.s.c + 1;
        if (colCount <= 0) return;

        const firstColumn = this._getOrCreateColumn(0);
        if (!firstColumn.totalsRowFunction && !firstColumn.totalsRowLabel) {
            firstColumn.totalsRowLabel = '\u603B\u8BA1';
        }

        for (let colIndex = 0; colIndex < colCount; colIndex++) {
            const column = this._getOrCreateColumn(colIndex);
            if (column.totalsRowFunction || column.totalsRowLabel) continue;

            if (this._isNumericTotalsCandidate(colIndex)) {
                column.totalsRowFunction = 'sum';
            }
        }
    }

    _applyTotalsCell(colIndex, options = {}) {
        if (!this.showTotalsRow || !this._ref) return;

        const force = options.force === true;
        const totalsRow = this.totalsRowIndex;
        if (totalsRow < 0) return;

        const column = this._getOrCreateColumn(colIndex);
        const rawFunction = (column.totalsRowFunction ?? '').toString().trim().toLowerCase();
        if (rawFunction === 'custom') return;

        const formula = this._buildTotalsFormula(colIndex, column.totalsRowFunction);
        const label = (column.totalsRowLabel ?? '').toString();
        const targetValue = formula || label;

        if (!force && !targetValue) return;

        const cell = this.sheet.getCell(totalsRow, this._ref.s.c + colIndex);
        if (cell.editVal !== targetValue) {
            cell.editVal = targetValue;
        }
    }

    /** @param {{initializeDefaults?: boolean, force?: boolean}} options */
    applyTotalsRow(options = {}) {
        if (!this.showTotalsRow || !this._ref) return;

        const colCount = this._ref.e.c - this._ref.s.c + 1;
        if (colCount <= 0) return;

        while (this.columns.length < colCount) {
            this.columns.push({
                id: this.columns.length + 1,
                name: `\u5217${this.columns.length + 1}`
            });
        }

        if (options.initializeDefaults) {
            this._ensureTotalsDefaults();
        }

        for (let colIndex = 0; colIndex < colCount; colIndex++) {
            this._applyTotalsCell(colIndex, { force: options.force === true });
        }
    }

    /** @param {number} colIndex @param {string|null} funcType */
    setTotalsFunction(colIndex, funcType) {
        if (!this._ref) return;

        const colCount = this._ref.e.c - this._ref.s.c + 1;
        if (colIndex < 0 || colIndex >= colCount) {
            throw new Error('Column index is out of table range');
        }

        const raw = funcType == null ? '' : String(funcType).trim();
        const normalized = this._normalizeTotalsFunction(raw);
        const lowered = raw.toLowerCase();
        if (!normalized && raw && lowered !== 'none' && lowered !== 'null') {
            throw new Error(`Unsupported totals function: ${funcType}`);
        }

        const column = this._getOrCreateColumn(colIndex);
        column.totalsRowFunction = normalized;

        if (normalized) {
            column.totalsRowLabel = null;
        } else if (colIndex === 0 && !column.totalsRowLabel) {
            column.totalsRowLabel = '\u603B\u8BA1';
        }

        if (!this.showTotalsRow) {
            this.showTotalsRow = true;
            return;
        }

        this._applyTotalsCell(colIndex, { force: true });
    }

    /** Convert table to plain range. */
    convertToRange() {
        // 清除筛选
        if (this.autoFilterEnabled) {
            this.sheet.AutoFilter.unregisterTableScope(this.id);
        }

        // 从工作表中移除
        this.sheet.Table.remove(this.id);
    }

    // ===================== Serialization =====================

    /** @returns {Object} */
    toJSON() {
        return {
            id: this.id,
            name: this.name,
            displayName: this.displayName,
            ref: this.ref,
            columns: this.columns,
            styleId: this.styleId,
            showHeaderRow: this.showHeaderRow,
            showTotalsRow: this.showTotalsRow,
            showFirstColumn: this.showFirstColumn,
            showLastColumn: this.showLastColumn,
            showRowStripes: this.showRowStripes,
            showColumnStripes: this.showColumnStripes,
            autoFilterEnabled: this.autoFilterEnabled,
            headerRowDxfId: this.headerRowDxfId,
            dataDxfId: this.dataDxfId,
            headerRowBorderDxfId: this.headerRowBorderDxfId,
            tableBorderDxfId: this.tableBorderDxfId,
            totalsRowBorderDxfId: this.totalsRowBorderDxfId
        };
    }

    /** @param {Sheet} sheet @param {Object} json @returns {TableItem} */
    static fromJSON(sheet, json) {
        return new TableItem(sheet, json);
    }
}
