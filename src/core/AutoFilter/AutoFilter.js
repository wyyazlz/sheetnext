/**
 * 自动筛选模块
 * @title 🔍 自动筛选
 * 负责管理工作表筛选状态（支持 Sheet 级 + Table 级独立筛选）
 * @class
 */

const DEFAULT_SCOPE_ID = '__sheet__';
const TABLE_SCOPE_PREFIX = 'table:';

function cloneRange(range) {
    if (!range?.s || !range?.e) return null;
    return {
        s: { r: range.s.r, c: range.s.c },
        e: { r: range.e.r, c: range.e.c }
    };
}

function normalizeRange(range) {
    if (!range?.s || !range?.e) return null;
    const sr = Math.min(range.s.r, range.e.r);
    const er = Math.max(range.s.r, range.e.r);
    const sc = Math.min(range.s.c, range.e.c);
    const ec = Math.max(range.s.c, range.e.c);
    return { s: { r: sr, c: sc }, e: { r: er, c: ec } };
}

function isRangeEqual(a, b) {
    if (!a && !b) return true;
    if (!a || !b) return false;
    return a.s.r === b.s.r && a.s.c === b.s.c && a.e.r === b.e.r && a.e.c === b.e.c;
}

export default class AutoFilter {
    /**
     * @param {Sheet} sheet - 所属工作表
     */
    constructor(sheet) {
        /**
         * 所属工作表
         * @type {Sheet}
         */
        this.sheet = sheet;
        /**
         * SheetNext 主实例
         * @type {Object}
         */
        this._SN = sheet.SN;

        this._defaultScopeId = DEFAULT_SCOPE_ID;
        this._activeScopeId = DEFAULT_SCOPE_ID;

        /**
         * 多筛选作用域：默认包含 __sheet__
         * @type {Map<string, Object>}
         */
        this._scopes = new Map();
        this._scopes.set(DEFAULT_SCOPE_ID, this._createScope(DEFAULT_SCOPE_ID, {
            ownerType: 'sheet',
            ownerId: null
        }));

        /**
         * 所有筛选作用域合并后的隐藏行
         * @type {Set<number>}
         */
        this._allHiddenRows = new Set();
    }

    _createScope(id, options = {}) {
        return {
            id,
            ownerType: options.ownerType || 'sheet',
            ownerId: options.ownerId ?? null,
            ref: null,
            range: null,
            columns: new Map(),
            hiddenRows: new Set(),
            sortState: null
        };
    }

    _resolveScopeId(scopeId = null) {
        if (scopeId === null || scopeId === undefined) {
            return this._defaultScopeId;
        }
        if (scopeId === 'sheet') return this._defaultScopeId;
        return scopeId;
    }

    _getScope(scopeId = null, create = false, options = {}) {
        const resolvedId = this._resolveScopeId(scopeId);
        if (!this._scopes.has(resolvedId) && create) {
            this._scopes.set(resolvedId, this._createScope(resolvedId, options));
        }
        return this._scopes.get(resolvedId) || null;
    }

    _getDefaultScope() {
        return this._getScope(this._defaultScopeId, true, { ownerType: 'sheet', ownerId: null });
    }

    _setScopeRange(scope, range, options = {}) {
        const normalized = normalizeRange(range);
        scope.range = normalized;
        scope.ref = normalized ? this._SN.Utils.rangeNumToStr(normalized) : null;

        if (!normalized) {
            scope.columns.clear();
            scope.sortState = null;
            scope.hiddenRows.clear();
            return;
        }

        if (options.keepFilters !== true) {
            scope.columns.clear();
            scope.sortState = null;
            scope.hiddenRows.clear();
            return;
        }

        // 范围变化时，清理越界列筛选条件
        Array.from(scope.columns.keys()).forEach(colIndex => {
            if (colIndex < normalized.s.c || colIndex > normalized.e.c) {
                scope.columns.delete(colIndex);
            }
        });

        if (scope.sortState && (scope.sortState.colIndex < normalized.s.c || scope.sortState.colIndex > normalized.e.c)) {
            scope.sortState = null;
        }

        scope.hiddenRows.clear();
    }

    _collectScopeHiddenRows(scope) {
        if (!scope?.range) return new Set();
        if (scope.columns.size === 0) return new Set();

        const hiddenRows = new Set();
        const range = scope.range;
        const sheet = this.sheet;

        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            let shouldHide = false;
            for (const [colIndex, filter] of scope.columns) {
                const cell = sheet.getCell(r, colIndex);
                const val = cell.editVal;
                if (!this._matchFilter(val, filter)) {
                    shouldHide = true;
                    break;
                }
            }
            if (shouldHide) hiddenRows.add(r);
        }

        return hiddenRows;
    }

    _syncCombinedHiddenRows() {
        const nextHiddenRows = new Set();
        this._scopes.forEach(scope => {
            if (!scope?.range || scope.hiddenRows.size === 0) return;
            scope.hiddenRows.forEach(r => nextHiddenRows.add(r));
        });

        const prevHiddenRows = this._allHiddenRows;
        const hiddenRowsAdded = [];
        const hiddenRowsRemoved = [];

        prevHiddenRows.forEach(r => {
            if (!nextHiddenRows.has(r)) {
                const row = this.sheet.rows[r];
                if (row) row._filterHidden = false;
                hiddenRowsRemoved.push(r);
            }
        });

        nextHiddenRows.forEach(r => {
            if (!prevHiddenRows.has(r)) {
                const row = this.sheet.getRow(r);
                row._filterHidden = true;
                hiddenRowsAdded.push(r);
            }
        });

        this._allHiddenRows = nextHiddenRows;
        return { hiddenRowsAdded, hiddenRowsRemoved };
    }

    _normalizeScopeForEvents(scopeId, scope) {
        return {
            scopeId,
            ownerType: scope?.ownerType || 'sheet',
            ownerId: scope?.ownerId ?? null
        };
    }

    // ===================== 对外：默认 scope 兼容访问 =====================

    /**
     * 获取/设置默认（sheet）筛选范围
     * @type {string|null}
     */
    get ref() {
        return this._getDefaultScope().ref;
    }

    set ref(val) {
        if (!val) {
            this.clear(this._defaultScopeId);
            return;
        }
        const range = this.sheet.rangeStrToNum(val);
        this.setRange(range, this._defaultScopeId);
    }

    /**
     * 兼容旧代码：直接读写 _ref（默认作用域）
     */
    get _ref() {
        return this.ref;
    }

    set _ref(val) {
        const scope = this._getDefaultScope();
        if (!val) {
            scope.ref = null;
            scope.range = null;
        } else {
            scope.ref = val;
            scope.range = this.sheet.rangeStrToNum(val);
        }
    }

    /**
     * 默认（sheet）作用域范围
     * @type {{s: {r: number, c: number}, e: {r: number, c: number}}|null}
     */
    get range() {
        return this._getDefaultScope().range;
    }

    /**
     * 兼容旧代码：直接读写 _range（默认作用域）
     */
    get _range() {
        return this.range;
    }

    set _range(val) {
        const scope = this._getDefaultScope();
        const normalized = normalizeRange(val);
        scope.range = normalized;
        scope.ref = normalized ? this._SN.Utils.rangeNumToStr(normalized) : null;
    }

    /**
     * 默认（sheet）作用域筛选条件
     * @type {Map<number, Object>}
     */
    get columns() {
        return this._getDefaultScope().columns;
    }

    /**
     * 默认（sheet）作用域排序状态
     * @type {{colIndex: number, order: 'asc'|'desc'}|null}
     */
    get sortState() {
        return this._getDefaultScope().sortState;
    }

    set sortState(val) {
        this._getDefaultScope().sortState = val || null;
    }

    /**
     * 默认（sheet）作用域是否启用
     * @type {boolean}
     */
    get enabled() {
        const scope = this._getDefaultScope();
        return !!scope?.range;
    }

    /**
     * 默认（sheet）作用域表头行
     * @type {number}
     */
    get headerRow() {
        return this._getDefaultScope()?.range?.s?.r ?? 0;
    }

    /**
     * 默认（sheet）作用域是否有生效筛选
     * @type {boolean}
     */
    get hasActiveFilters() {
        return this._getDefaultScope().columns.size > 0;
    }

    /**
     * 所有筛选隐藏行（只读，供公式等内部逻辑复用）
     * @type {Set<number>}
     */
    get _hiddenRows() {
        return this._allHiddenRows;
    }

    // ===================== scope 管理 =====================

    /**
     * 当前激活作用域 ID
     * @type {string}
     */
    get activeScopeId() {
        return this._activeScopeId;
    }

    /** @param {string|null} scopeId @returns {string} Set active scope id when scope exists. */
    setActiveScope(scopeId) {
        const resolvedId = this._resolveScopeId(scopeId);
        if (this._scopes.has(resolvedId)) {
            this._activeScopeId = resolvedId;
        }
        return this._activeScopeId;
    }

    /** @param {string} tableId @returns {string} Build table scope id from table id. */
    getTableScopeId(tableId) {
        return `${TABLE_SCOPE_PREFIX}${tableId}`;
    }

    /** @param {string|null} [scopeId=null] @returns {boolean} Check whether scope range is enabled. */
    isScopeEnabled(scopeId = null) {
        const scope = this._getScope(scopeId);
        return !!scope?.range;
    }

    /** @param {string|null} [scopeId=null] @returns {boolean} Check whether scope has active column filters. */
    hasActiveFiltersInScope(scopeId = null) {
        const scope = this._getScope(scopeId);
        return !!scope && scope.columns.size > 0;
    }

    /** @returns {boolean} Check whether any scope has active filters. */
    hasAnyActiveFilters() {
        for (const scope of this._scopes.values()) {
            if (scope.columns.size > 0) return true;
        }
        return false;
    }

    /** @param {string|null} [scopeId=null] @returns {{s:{r:number,c:number},e:{r:number,c:number}}|null} Get range of scope. */
    getScopeRange(scopeId = null) {
        const scope = this._getScope(scopeId);
        return scope?.range || null;
    }

    /** @param {string|null} [scopeId=null] @returns {number} Get header row index of scope. */
    getScopeHeaderRow(scopeId = null) {
        const scope = this._getScope(scopeId);
        return scope?.range?.s?.r ?? 0;
    }

    /** @param {number} colIndex @param {string|null} [scopeId=null] @returns {Object|null} Get filter config for column in scope. */
    getColumnFilter(colIndex, scopeId = null) {
        const scope = this._getScope(scopeId);
        return scope?.columns?.get(colIndex) || null;
    }

    /** @returns {Array<{scopeId:string,ownerType:string,ownerId:string|null,range:{s:{r:number,c:number},e:{r:number,c:number}},headerRow:number}>} List enabled scopes sorted by range. */
    getEnabledScopes() {
        return Array.from(this._scopes.values())
            .filter(scope => !!scope?.range)
            .sort((a, b) => {
                if (a.range.s.r !== b.range.s.r) return a.range.s.r - b.range.s.r;
                if (a.range.s.c !== b.range.s.c) return a.range.s.c - b.range.s.c;
                return a.id.localeCompare(b.id);
            })
            .map(scope => ({
                scopeId: scope.id,
                ownerType: scope.ownerType,
                ownerId: scope.ownerId,
                range: cloneRange(scope.range),
                headerRow: scope.range.s.r
            }));
    }

    _getTableFilterRange(table) {
        if (!table?.range) return null;
        return {
            s: { r: table.range.s.r, c: table.range.s.c },
            e: { r: table.showTotalsRow ? table.range.e.r - 1 : table.range.e.r, c: table.range.e.c }
        };
    }

    /** @param {Object} table @param {Object} [options={}] @returns {string|null} Register or update filter scope for table. */
    registerTableScope(table, options = {}) {
        if (!table?.id) return null;

        const scopeId = this.getTableScopeId(table.id);
        const scope = this._getScope(scopeId, true, {
            ownerType: 'table',
            ownerId: table.id
        });

        scope.ownerType = 'table';
        scope.ownerId = table.id;

        const oldRange = cloneRange(scope.range);
        const range = normalizeRange(options.range || this._getTableFilterRange(table));
        if (!range) {
            this.unregisterTableScope(table.id, options);
            return null;
        }

        const keepFilters = options.keepFilters !== false;
        this._setScopeRange(scope, range, { keepFilters });

        if (options.autoFilterXml) {
            const parsed = this._parseAutoFilterXml(options.autoFilterXml, range);
            this._applyParsedScope(scope, parsed, {
                restoreHiddenRowsFromSheet: options.restoreHiddenRowsFromSheet === true
            });
        }

        if (scope.columns.size > 0) {
            this._applyFilters(scopeId, {
                silent: options.silent === true,
                emitEvents: false,
                recordChange: false,
                recalculate: options.silent !== true
            });
        } else {
            scope.hiddenRows.clear();
            this._syncCombinedHiddenRows();
            if (options.silent !== true && !isRangeEqual(oldRange, scope.range)) {
                this._SN._r();
            }
        }

        return scopeId;
    }

    /** @param {string} tableId @param {{silent?:boolean}} [options={}] @returns {void} Unregister table scope and refresh view state. */
    unregisterTableScope(tableId, options = {}) {
        const scopeId = this.getTableScopeId(tableId);
        const scope = this._getScope(scopeId);
        if (!scope) return;

        const hadHiddenRows = scope.hiddenRows.size > 0;
        this._scopes.delete(scopeId);

        if (this._activeScopeId === scopeId) {
            this._activeScopeId = this._defaultScopeId;
        }

        if (hadHiddenRows) {
            this._syncCombinedHiddenRows();
        }

        if (options.silent !== true) {
            this._SN._r();
        }
    }

    // ===================== 默认筛选范围管理 =====================

    /**
     * 设置筛选范围（默认 sheet scope；可指定 scopeId）
     * @param {{s: {r: number, c: number}, e: {r: number, c: number}}} range
     * @param {string} [scopeId]
     */
    setRange(range, scopeId = this._defaultScopeId) {
        const resolvedScopeId = this._resolveScopeId(scopeId);
        const scope = this._getScope(resolvedScopeId, true, {
            ownerType: resolvedScopeId === this._defaultScopeId ? 'sheet' : 'custom',
            ownerId: null
        });

        let nextRange = range;
        if (!nextRange) {
            nextRange = this._detectDataRange();
        }
        nextRange = normalizeRange(nextRange);
        if (!nextRange) return;

        const eventBase = this._normalizeScopeForEvents(resolvedScopeId, scope);

        // 通用保护事件
        if (this._SN.Event.hasListeners('beforeAutoFilter')) {
            const protectEvent = this._SN.Event.emit('beforeAutoFilter', {
                sheet: this.sheet,
                action: 'setRange',
                range: nextRange,
                ...eventBase
            });
            if (protectEvent.canceled) return;
        }

        const beforeEvent = this._SN.Event.emit('beforeAutoFilterSetRange', {
            sheet: this.sheet,
            range: nextRange,
            ...eventBase
        });
        if (beforeEvent.canceled) return;

        return this._SN._withOperation('autoFilterSetRange', { sheet: this.sheet, range: nextRange, scopeId: resolvedScopeId }, () => {
            this._setScopeRange(scope, nextRange, { keepFilters: false });
            this._syncCombinedHiddenRows();

            this._SN._recordChange({
                type: 'autoFilter',
                action: 'setRange',
                sheet: this.sheet,
                range: nextRange,
                scopeId: resolvedScopeId
            });
            this._SN._r();
            this._SN.Event.emit('afterAutoFilterSetRange', {
                sheet: this.sheet,
                range: nextRange,
                ...eventBase
            });
        });
    }

    _detectDataRange() {
        const sheet = this.sheet;
        const areas = sheet.activeAreas;

        // 如果有选区且不是单个单元格，使用选区
        if (areas.length > 0) {
            const area = areas[0];
            if (area.s.r !== area.e.r || area.s.c !== area.e.c) {
                return area;
            }
        }

        // 否则从当前单元格扩展到连续数据区域
        const startCell = sheet.activeCell;
        let minR = startCell.r, maxR = startCell.r;
        let minC = startCell.c, maxC = startCell.c;

        while (minR > 0 && this._hasData(minR - 1, startCell.c)) minR--;
        while (maxR < sheet.rowCount - 1 && this._hasData(maxR + 1, startCell.c)) maxR++;
        while (minC > 0 && this._hasData(startCell.r, minC - 1)) minC--;
        while (maxC < sheet.colCount - 1 && this._hasData(startCell.r, maxC + 1)) maxC++;

        // 确保至少有两行（标题+数据）
        if (maxR === minR) maxR = Math.min(minR + 100, sheet.rowCount - 1);

        return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC } };
    }

    _hasData(r, c) {
        const cell = this.sheet.getCell(r, c);
        return cell.editVal !== undefined && cell.editVal !== null && cell.editVal !== '';
    }

    /**
     * 判断某列是否在筛选范围内
     * @param {number} colIndex
     * @param {string} [scopeId]
     */
    isColumnInRange(colIndex, scopeId = null) {
        const scope = this._getScope(scopeId);
        if (!scope?.range) return false;
        return colIndex >= scope.range.s.c && colIndex <= scope.range.e.c;
    }

    /**
     * 判断某行是否是表头行
     * @param {number} rowIndex
     * @param {string} [scopeId]
     */
    isHeaderRow(rowIndex, scopeId = null) {
        const scope = this._getScope(scopeId);
        return !!scope?.range && rowIndex === scope.range.s.r;
    }

    /**
     * 获取某列的所有唯一值（用于筛选面板）
     * @param {number} colIndex
     * @param {string} [scopeId]
     * @returns {Array<{value: any, text: string, count: number, isEmpty: boolean}>}
     */
    getColumnValues(colIndex, scopeId = null) {
        const scope = this._getScope(scopeId);
        if (!scope?.range) return [];

        const values = new Map();
        const sheet = this.sheet;

        for (let r = scope.range.s.r + 1; r <= scope.range.e.r; r++) {
            const cell = sheet.getCell(r, colIndex);
            const val = cell.editVal;
            const displayVal = cell.showVal ?? val ?? '';
            const key = String(val ?? '');

            if (values.has(key)) {
                values.get(key).count++;
            } else {
                values.set(key, {
                    value: val,
                    text: displayVal === '' ? this._SN.t('core.autoFilter.autoFilter.blank') : String(displayVal),
                    count: 1,
                    isEmpty: val === '' || val === null || val === undefined
                });
            }
        }

        return Array.from(values.values()).sort((a, b) => {
            if (a.isEmpty && !b.isEmpty) return 1;
            if (!a.isEmpty && b.isEmpty) return -1;

            const aNum = typeof a.value === 'number';
            const bNum = typeof b.value === 'number';
            if (aNum && !bNum) return -1;
            if (!aNum && bNum) return 1;
            if (aNum && bNum) return a.value - b.value;

            return String(a.text).localeCompare(String(b.text));
        });
    }

    /**
     * 设置列筛选条件
     * @param {number} colIndex
     * @param {Object} filter
     * @param {string} [scopeId]
     */
    setColumnFilter(colIndex, filter, scopeId = null) {
        const resolvedScopeId = this._resolveScopeId(scopeId);
        const scope = this._getScope(resolvedScopeId);
        if (!scope?.range) return;

        const eventBase = this._normalizeScopeForEvents(resolvedScopeId, scope);

        if (this._SN.Event.hasListeners('beforeAutoFilter')) {
            const protectEvent = this._SN.Event.emit('beforeAutoFilter', {
                sheet: this.sheet,
                action: 'setColumnFilter',
                colIndex,
                filter,
                ...eventBase
            });
            if (protectEvent.canceled) return;
        }

        const prevFilter = scope.columns.get(colIndex);
        const beforeEvent = this._SN.Event.emit('beforeAutoFilterSetColumnFilter', {
            sheet: this.sheet,
            colIndex,
            filter,
            prevFilter,
            ...eventBase
        });
        if (beforeEvent.canceled) return;

        return this._SN._withOperation('autoFilterSetColumnFilter', { sheet: this.sheet, colIndex, filter, scopeId: resolvedScopeId }, () => {
            if (!filter || (filter.type === 'values' && (!filter.values || filter.values.length === 0))) {
                scope.columns.delete(colIndex);
            } else {
                scope.columns.set(colIndex, filter);
            }

            this._SN._recordChange({
                type: 'autoFilter',
                action: 'setColumnFilter',
                sheet: this.sheet,
                colIndex,
                filter,
                prevFilter,
                scopeId: resolvedScopeId
            });

            this._applyFilters(resolvedScopeId);

            this._SN.Event.emit('afterAutoFilterSetColumnFilter', {
                sheet: this.sheet,
                colIndex,
                filter,
                prevFilter,
                ...eventBase
            });
        });
    }

    /**
     * 清除某列筛选
     * @param {number} colIndex
     * @param {string} [scopeId]
     */
    clearColumnFilter(colIndex, scopeId = null) {
        const resolvedScopeId = this._resolveScopeId(scopeId);
        const scope = this._getScope(resolvedScopeId);
        if (!scope?.range) return;

        const eventBase = this._normalizeScopeForEvents(resolvedScopeId, scope);

        if (this._SN.Event.hasListeners('beforeAutoFilter')) {
            const protectEvent = this._SN.Event.emit('beforeAutoFilter', {
                sheet: this.sheet,
                action: 'clearColumnFilter',
                colIndex,
                ...eventBase
            });
            if (protectEvent.canceled) return;
        }

        const prevFilter = scope.columns.get(colIndex);
        const beforeEvent = this._SN.Event.emit('beforeAutoFilterClearColumnFilter', {
            sheet: this.sheet,
            colIndex,
            prevFilter,
            ...eventBase
        });
        if (beforeEvent.canceled) return;

        return this._SN._withOperation('autoFilterClearColumnFilter', { sheet: this.sheet, colIndex, scopeId: resolvedScopeId }, () => {
            scope.columns.delete(colIndex);

            this._SN._recordChange({
                type: 'autoFilter',
                action: 'clearColumnFilter',
                sheet: this.sheet,
                colIndex,
                prevFilter,
                scopeId: resolvedScopeId
            });

            this._applyFilters(resolvedScopeId);

            this._SN.Event.emit('afterAutoFilterClearColumnFilter', {
                sheet: this.sheet,
                colIndex,
                prevFilter,
                ...eventBase
            });
        });
    }

    /**
     * 清除筛选条件（保留范围）
     * @param {string} [scopeId]
     */
    clearAllFilters(scopeId = null) {
        const resolvedScopeId = this._resolveScopeId(scopeId);
        const scope = this._getScope(resolvedScopeId);
        if (!scope?.range) return;

        const eventBase = this._normalizeScopeForEvents(resolvedScopeId, scope);

        if (this._SN.Event.hasListeners('beforeAutoFilter')) {
            const protectEvent = this._SN.Event.emit('beforeAutoFilter', {
                sheet: this.sheet,
                action: 'clearAllFilters',
                ...eventBase
            });
            if (protectEvent.canceled) return;
        }

        const prevFilters = new Map(scope.columns);
        const beforeEvent = this._SN.Event.emit('beforeAutoFilterClearAllFilters', {
            sheet: this.sheet,
            filters: prevFilters,
            ...eventBase
        });
        if (beforeEvent.canceled) return;

        return this._SN._withOperation('autoFilterClearAllFilters', { sheet: this.sheet, scopeId: resolvedScopeId }, () => {
            scope.columns.clear();

            this._SN._recordChange({
                type: 'autoFilter',
                action: 'clearAllFilters',
                sheet: this.sheet,
                filters: prevFilters,
                scopeId: resolvedScopeId
            });

            this._applyFilters(resolvedScopeId);

            this._SN.Event.emit('afterAutoFilterClearAllFilters', {
                sheet: this.sheet,
                ...eventBase
            });
        });
    }

    /**
     * 完全清除筛选（包含范围）
     * @param {string} [scopeId]
     */
    clear(scopeId = this._defaultScopeId) {
        const resolvedScopeId = this._resolveScopeId(scopeId);
        const scope = this._getScope(resolvedScopeId);
        if (!scope) return;

        const eventBase = this._normalizeScopeForEvents(resolvedScopeId, scope);

        if (this._SN.Event.hasListeners('beforeAutoFilter')) {
            const protectEvent = this._SN.Event.emit('beforeAutoFilter', {
                sheet: this.sheet,
                action: 'clear',
                ...eventBase
            });
            if (protectEvent.canceled) return;
        }

        const prevRef = scope.ref;
        const prevRange = cloneRange(scope.range);
        const prevFilters = new Map(scope.columns);

        const beforeEvent = this._SN.Event.emit('beforeAutoFilterClear', {
            sheet: this.sheet,
            ref: prevRef,
            range: prevRange,
            ...eventBase
        });
        if (beforeEvent.canceled) return;

        return this._SN._withOperation('autoFilterClear', { sheet: this.sheet, scopeId: resolvedScopeId }, () => {
            scope.ref = null;
            scope.range = null;
            scope.columns.clear();
            scope.sortState = null;
            scope.hiddenRows.clear();

            const globalDelta = this._syncCombinedHiddenRows();

            this._SN._recordChange({
                type: 'autoFilter',
                action: 'clear',
                sheet: this.sheet,
                ref: prevRef,
                range: prevRange,
                filters: prevFilters,
                hiddenRowsAdded: globalDelta.hiddenRowsAdded,
                hiddenRowsRemoved: globalDelta.hiddenRowsRemoved,
                scopeId: resolvedScopeId
            });

            if (this._SN.DependencyGraph && this._SN.calcMode !== 'manual') {
                this._SN.DependencyGraph.recalculateVolatile();
            }

            this._SN._r();
            this._SN.Event.emit('afterAutoFilterClear', {
                sheet: this.sheet,
                ref: prevRef,
                range: prevRange,
                ...eventBase
            });
        });
    }

    _applyFilters(scopeId = null, options = {}) {
        const resolvedScopeId = this._resolveScopeId(scopeId);
        const scope = this._getScope(resolvedScopeId);
        if (!scope?.range) return;

        const emitEvents = options.emitEvents !== false;
        const recordChange = options.recordChange !== false;
        const silent = options.silent === true;
        const recalculate = options.recalculate !== false;
        const eventBase = this._normalizeScopeForEvents(resolvedScopeId, scope);

        const runApply = () => {
            const prevHiddenRows = scope.hiddenRows;
            const newHiddenRows = this._collectScopeHiddenRows(scope);
            const hiddenRowsAdded = [];
            const hiddenRowsRemoved = [];

            prevHiddenRows.forEach(r => {
                if (!newHiddenRows.has(r)) hiddenRowsRemoved.push(r);
            });
            newHiddenRows.forEach(r => {
                if (!prevHiddenRows.has(r)) hiddenRowsAdded.push(r);
            });

            scope.hiddenRows = newHiddenRows;
            const globalDelta = this._syncCombinedHiddenRows();

            if (recordChange) {
                this._SN._recordChange({
                    type: 'autoFilter',
                    action: 'apply',
                    sheet: this.sheet,
                    hiddenRowsAdded,
                    hiddenRowsRemoved,
                    allHiddenRowsAdded: globalDelta.hiddenRowsAdded,
                    allHiddenRowsRemoved: globalDelta.hiddenRowsRemoved,
                    scopeId: resolvedScopeId
                });
            }

            if (recalculate && this._SN.DependencyGraph && this._SN.calcMode !== 'manual') {
                this._SN.DependencyGraph.recalculateVolatile();
            }

            if (!silent) {
                this._SN._r();
            }

            if (emitEvents) {
                this._SN.Event.emit('afterAutoFilterApply', {
                    sheet: this.sheet,
                    range: scope.range,
                    hiddenRowsAdded,
                    hiddenRowsRemoved,
                    allHiddenRowsAdded: globalDelta.hiddenRowsAdded,
                    allHiddenRowsRemoved: globalDelta.hiddenRowsRemoved,
                    ...eventBase
                });
            }
        };

        if (emitEvents) {
            const beforeEvent = this._SN.Event.emit('beforeAutoFilterApply', {
                sheet: this.sheet,
                range: scope.range,
                columns: scope.columns,
                ...eventBase
            });
            if (beforeEvent.canceled) return;
        }

        if (recordChange) {
            return this._SN._withOperation(
                'autoFilterApply',
                { sheet: this.sheet, range: scope.range, scopeId: resolvedScopeId },
                runApply,
                { joinParent: false }
            );
        }

        runApply();
    }

    _matchFilter(val, filter) {
        if (filter.type === 'values') {
            const strVal = String(val ?? '');
            return filter.values.some(v => String(v ?? '') === strVal);
        }
        if (filter.type === 'custom') {
            return this._matchCustomFilter(val, filter.filters, filter.and);
        }
        return true;
    }

    _matchCustomFilter(val, filters, useAnd = true) {
        if (!filters || filters.length === 0) return true;

        const results = filters.map(f => this._matchSingleCondition(val, f));
        return useAnd ? results.every(Boolean) : results.some(Boolean);
    }

    _matchSingleCondition(val, condition) {
        const { operator, value } = condition;
        const numVal = typeof val === 'number' ? val : parseFloat(val);
        const numCondVal = parseFloat(value);
        const strVal = String(val ?? '').toLowerCase();
        const strCondVal = String(value ?? '').toLowerCase();

        switch (operator) {
            case 'equal': return strVal === strCondVal;
            case 'notEqual': return strVal !== strCondVal;
            case 'greaterThan': return numVal > numCondVal;
            case 'greaterThanOrEqual': return numVal >= numCondVal;
            case 'lessThan': return numVal < numCondVal;
            case 'lessThanOrEqual': return numVal <= numCondVal;
            case 'beginsWith': return strVal.startsWith(strCondVal);
            case 'endsWith': return strVal.endsWith(strCondVal);
            case 'contains': return strVal.includes(strCondVal);
            case 'notContains': return !strVal.includes(strCondVal);
            default: return true;
        }
    }

    /**
     * 判断列是否有筛选
     * @param {number} colIndex
     * @param {string} [scopeId]
     */
    hasFilter(colIndex, scopeId = null) {
        const scope = this._getScope(scopeId);
        return !!scope && scope.columns.has(colIndex);
    }

    /**
     * 设置排序状态
     * @param {number} colIndex
     * @param {'asc'|'desc'} order
     * @param {string} [scopeId]
     */
    setSortState(colIndex, order, scopeId = null) {
        const scope = this._getScope(scopeId);
        if (!scope) return;
        scope.sortState = { colIndex, order };
    }

    /**
     * 获取排序状态
     * @param {number} colIndex
     * @param {string} [scopeId]
     */
    getSortOrder(colIndex, scopeId = null) {
        const scope = this._getScope(scopeId);
        if (scope?.sortState && scope.sortState.colIndex === colIndex) {
            return scope.sortState.order;
        }
        return null;
    }

    /**
     * 行是否因筛选被隐藏
     * @param {number} rowIndex
     * @returns {boolean}
     */
    isRowFilteredHidden(rowIndex) {
        return this._allHiddenRows.has(rowIndex);
    }

    // ===================== XML 解析/构建 =====================

    _parseAutoFilterXml(autoFilterXml, fallbackRange = null) {
        if (!autoFilterXml) return null;

        const fallback = normalizeRange(fallbackRange);
        const refFromXml = autoFilterXml['_$ref'];
        const range = refFromXml
            ? normalizeRange(this.sheet.rangeStrToNum(refFromXml))
            : fallback;

        if (!range) return null;

        const result = {
            ref: refFromXml || this._SN.Utils.rangeNumToStr(range),
            range,
            columns: new Map(),
            sortState: null
        };

        const filterColumns = autoFilterXml.filterColumn;
        if (filterColumns) {
            const columns = Array.isArray(filterColumns) ? filterColumns : [filterColumns];
            columns.forEach(col => {
                const rawColId = parseInt(col['_$colId'], 10);
                if (!Number.isFinite(rawColId)) return;

                // 兼容：优先按 OOXML 规则（相对列），超界时回退为旧版绝对列
                const width = range.e.c - range.s.c + 1;
                let colId = rawColId;
                if (rawColId >= 0 && rawColId < width) {
                    colId = range.s.c + rawColId;
                }

                if (col.filters) {
                    const filters = col.filters.filter;
                    const values = [];
                    if (filters) {
                        const filterArr = Array.isArray(filters) ? filters : [filters];
                        filterArr.forEach(f => {
                            values.push(f['_$val']);
                        });
                    }
                    if (col.filters['_$blank'] === '1') {
                        values.push('');
                    }
                    result.columns.set(colId, { type: 'values', values });
                }

                if (col.customFilters) {
                    const customFilters = col.customFilters.customFilter;
                    const filters = [];
                    if (customFilters) {
                        const filterArr = Array.isArray(customFilters) ? customFilters : [customFilters];
                        filterArr.forEach(f => {
                            filters.push({
                                operator: f['_$operator'] || 'equal',
                                value: f['_$val']
                            });
                        });
                    }
                    const and = col.customFilters['_$and'] === '1';
                    result.columns.set(colId, { type: 'custom', filters, and });
                }
            });
        }

        const sortStateXml = autoFilterXml.sortState;
        if (sortStateXml) {
            const sortCondition = sortStateXml.sortCondition;
            if (sortCondition) {
                const condition = Array.isArray(sortCondition) ? sortCondition[0] : sortCondition;
                const condRef = condition['_$ref'];
                if (condRef) {
                    const colMatch = String(condRef).toUpperCase().match(/\$?([A-Z]+)/);
                    if (colMatch) {
                        const colIndex = this._SN.Utils.charToNum(colMatch[1]);
                        const descending = condition['_$descending'] === '1';
                        result.sortState = {
                            colIndex,
                            order: descending ? 'desc' : 'asc',
                            ref: sortStateXml['_$ref']
                        };
                    }
                }
            }
        }

        return result;
    }

    _applyParsedScope(scope, parsed, options = {}) {
        if (!scope) return;

        if (!parsed?.range) {
            scope.ref = null;
            scope.range = null;
            scope.columns.clear();
            scope.sortState = null;
            scope.hiddenRows.clear();
            this._syncCombinedHiddenRows();
            return;
        }

        scope.ref = parsed.ref;
        scope.range = parsed.range;
        scope.columns = parsed.columns;
        scope.sortState = parsed.sortState;
        scope.hiddenRows = new Set();

        if (options.restoreHiddenRowsFromSheet === true && scope.columns.size > 0) {
            for (let r = scope.range.s.r + 1; r <= scope.range.e.r; r++) {
                const row = this.sheet.rows[r];
                if (row && row._hidden) {
                    // 导入时将 XML hidden 行视为筛选隐藏，避免与手动隐藏混淆
                    row._hidden = false;
                    scope.hiddenRows.add(r);
                }
            }
        }

        this._syncCombinedHiddenRows();
    }

    /**
     * 从 worksheet AutoFilter 解析默认作用域
     * @param {Object} xmlObj
     */
    parse(xmlObj) {
        const autoFilterXml = xmlObj?.AutoFilter;
        if (!autoFilterXml) return;

        const scope = this._getDefaultScope();
        const parsed = this._parseAutoFilterXml(autoFilterXml, null);
        if (!parsed) return;

        this._applyParsedScope(scope, parsed, { restoreHiddenRowsFromSheet: true });
    }

    /**
     * 将指定作用域的筛选状态解析到内存（主要用于 table 导入）
     * @param {string} scopeId
     * @param {Object} autoFilterXml
     * @param {{range?: Object, restoreHiddenRowsFromSheet?: boolean, ownerType?: string, ownerId?: string}} [options]
     */
    parseScope(scopeId, autoFilterXml, options = {}) {
        const scope = this._getScope(scopeId, true, {
            ownerType: options.ownerType || 'custom',
            ownerId: options.ownerId ?? null
        });
        const parsed = this._parseAutoFilterXml(autoFilterXml, options.range || null);
        if (!parsed) return;

        this._applyParsedScope(scope, parsed, {
            restoreHiddenRowsFromSheet: options.restoreHiddenRowsFromSheet === true
        });
    }

    _buildScopeXml(scope) {
        if (!scope?.ref || !scope?.range) return null;

        const xml = { '_$ref': scope.ref };

        if (scope.columns.size > 0) {
            xml.filterColumn = [];

            for (const [colId, filter] of scope.columns) {
                const relColId = colId - scope.range.s.c;
                if (!Number.isFinite(relColId) || relColId < 0) continue;

                const colXml = { '_$colId': String(relColId) };

                if (filter.type === 'values') {
                    colXml.filters = {};
                    const hasBlank = filter.values.some(v => v === '' || v === null || v === undefined);
                    if (hasBlank) {
                        colXml.filters['_$blank'] = '1';
                    }
                    const nonBlankValues = filter.values.filter(v => v !== '' && v !== null && v !== undefined);
                    if (nonBlankValues.length > 0) {
                        colXml.filters.filter = nonBlankValues.map(v => ({ '_$val': String(v) }));
                    }
                } else if (filter.type === 'custom') {
                    colXml.customFilters = {};
                    if (filter.and) {
                        colXml.customFilters['_$and'] = '1';
                    }
                    colXml.customFilters.customFilter = filter.filters.map(f => ({
                        '_$operator': f.operator,
                        '_$val': String(f.value)
                    }));
                }

                xml.filterColumn.push(colXml);
            }
        }

        if (scope.sortState) {
            const range = scope.range;
            const sortRef = this._SN.Utils.rangeNumToStr({
                s: { r: range.s.r + 1, c: range.s.c },
                e: { r: range.e.r, c: range.e.c }
            });
            const colLetter = this._SN.Utils.numToChar(scope.sortState.colIndex);
            const condRef = `${colLetter}${range.s.r + 1}:${colLetter}${range.e.r + 1}`;

            xml.sortState = {
                '_$ref': sortRef,
                sortCondition: {
                    '_$ref': condRef
                }
            };

            if (scope.sortState.order === 'desc') {
                xml.sortState.sortCondition['_$descending'] = '1';
            }
        }

        return xml;
    }

    /**
     * 构建指定作用域 XML；默认导出 sheet 作用域
     * @param {string} [scopeId]
     * @returns {Object|null}
     */
    buildXml(scopeId = this._defaultScopeId) {
        const scope = this._getScope(scopeId);
        return this._buildScopeXml(scope);
    }
}
