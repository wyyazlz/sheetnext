import Row from '../Row/Row.js'
import Col from '../Col/Col.js'
import Cell from '../Cell/Cell.js'
import Comment from '../Comment/Comment.js'
import AutoFilter from '../AutoFilter/AutoFilter.js'
import Table from '../Table/Table.js'
import SheetProtection from './SheetProtection.js'

import * as Sort from './Sort.js';
import * as FlashFill from './FlashFill.js';
import * as TextToColumns from './TextToColumns.js';
import * as Outline from './Outline.js';
import * as Subtotal from './Subtotal.js';
import * as Consolidate from './Consolidate.js';
import * as DataValidation from './DataValidation.js';
import * as InsertTable from './InsertTable.js';
import * as AutoFit from './AutoFit.js';
import * as Merge from './Merge.js';
import * as RowColHandler from './RowColHandler.js';
import * as ViewCalculator from './ViewCalculator.js';
import * as Initializer from './Initializer.js';
import * as Clipboard from './Clipboard.js';
import * as CellControl from './CellControl.js';
import * as ClearRange from './ClearRange.js';
import * as UsedRange from './UsedRange.js';

/** @typedef {import('../Workbook/Workbook.js').default} SheetNext */

/** @class */
export default class Sheet {

    _name;
    _hidden;
    _activeCell;
    _defaultColWidth;
    _defaultRowHeight;

    /** @param {Object} meta @param {SheetNext} SN */
    constructor(meta, SN) {
        /** @type {import('../Workbook/Workbook.js').default} */
        this.SN = SN;
        /** @type {string} */
        this.rId = meta['_$r:id'];
        /** @type {Utils} */
        this.Utils = SN.Utils;
        /** @type {import('../Canvas/Canvas.js').default} */
        this.Canvas = SN.Canvas;
        /** @type {RangeNum[]} */
        this.merges = [];
        /** @type {Object} */
        this.vi = null;
        /** @type {Object[]} */
        this.views = [{ pane: {} }];
        /** @type {Row[]} */
        this.rows = [];
        /** @type {Col[]} */
        this.cols = [];
        /** @type {boolean} */
        this.initialized = false;
        /** @type {boolean} */
        this.showGridLines = true;
        /** @type {boolean} */
        this.showRowColHeaders = true;
        /** @type {boolean} */
        this.showPageBreaks = false;
        /** @type {Object} */
        this.outlinePr = {
            applyStyles: false,
            summaryBelow: true,
            summaryRight: true,
            showOutlineSymbols: true
        };
        /** @type {Object} */
        this.printSettings = null;
        /** @type {Object} */
        this.protection = new SheetProtection(this);
        this.Comment = new Comment(this);
        this.AutoFilter = new AutoFilter(this);
        this.Table = new Table(this);
        this.Drawing = null;
        this.Sparkline = null;
        this.CF = null;
        this.Slicer = null;
        this.PivotTable = null;

        this._xmlObj = null;
        this._name = meta['_$name'];
        this._hidden = meta['_$state'] == 'hidden';
        this._activeCell = { r: 0, c: 0 };
        this._defaultColWidth = 0;
        this._defaultRowHeight = 0;
        this._handleLayer = this.Canvas.handleLayer;
        this._brush = { keep: false, area: null };
        this._activeAreas = [];
        this._indexWidth = 50;
        this._headHeight = 23;
        this._styledRows = new Set();
        this._styledCols = new Set();
        this._rowVisibilityVersion = 0;
        this._rowMetricVersion = 0;
        this._colMetricVersion = 0;
        this._rowAxisMetricsCache = null;
        this._colAxisMetricsCache = null;
        this._frozenCols = 0;
        this._frozenRows = 0;
        this._freezeStartRow = 0;
        this._freezeStartCol = 0;
        this._virtualRowCount = 0;
        this._virtualColCount = 0;
        this._maxVirtualRowCount = 1048576;
        this._maxVirtualColCount = 16384;
        this._pendingXlsxRows = null;
        this._pendingXlsxRowMap = null;
        this._pendingXlsxRowIndexes = null;
        this._pendingXlsxRowMeta = null;
        this._pendingXlsxRowStateMap = null;
        this._zoom = 1;
        this._idCount = 0;
    }

    // ============================= Getters & Setters =============================

    /** @type {CellNum} */
    get activeCell() {
        return this._activeCell;
    }

    set activeCell(cell) {
        const oldCell = this._activeCell;
        // 触发 beforeSelectionChange 事件（可取消）
        const beforeEvent = this.SN.Event.emit('beforeSelectionChange', {
            sheet: this,
            oldCell,
            newCell: cell,
            type: 'activeCell'
        });
        if (beforeEvent.canceled) return;

        this._activeCell = cell;

        // 触发 afterSelectionChange 事件
        this.SN.Event.emit('afterSelectionChange', {
            sheet: this,
            oldCell,
            newCell: cell,
            type: 'activeCell'
        });
    }

    /** @type {RangeNum[]} */
    get activeAreas() {
        if (this._activeAreas.length > 0) return this._activeAreas;
        const cell = this.getCell(this.activeCell.r, this.activeCell.c);
        if (cell.isMerged) {
            this._activeAreas = [this.merges.find(m => m.s.r == cell.master.r && m.s.c == cell.master.c)]
        } else {
            this._activeAreas.push({ s: this.activeCell, e: this.activeCell });
        }
        return this._activeAreas
    }

    set activeAreas(arr) {
        const oldAreas = this._activeAreas;
        // 触发 beforeSelectionChange 事件（可取消）
        const beforeEvent = this.SN.Event.emit('beforeSelectionChange', {
            sheet: this,
            oldAreas,
            newAreas: arr,
            type: 'activeAreas'
        });
        if (beforeEvent.canceled) return;

        this._activeAreas = arr;

        // 触发 afterSelectionChange 事件
        this.SN.Event.emit('afterSelectionChange', {
            sheet: this,
            oldAreas,
            newAreas: arr,
            type: 'activeAreas'
        });
    }

    /** @type {number} 0.1-4 */
    get zoom() {
        return this._zoom ?? 1;
    }

    set zoom(value) {
        const num = Number(value);
        const next = Number.isFinite(num) ? Math.min(4, Math.max(0.1, num)) : 1;
        const oldZoom = this._zoom ?? 1;
        if (oldZoom === next) return;
        if (this.SN.Event.hasListeners('beforeZoomChange')) {
            const beforeEvent = this.SN.Event.emit('beforeZoomChange', { sheet: this, oldZoom, newZoom: next });
            if (beforeEvent.canceled) return;
        }
        this._zoom = next;
        if (this === this.SN.activeSheet) {
            this.SN.Canvas.applyZoom();
            this.SN._r();
            this.SN._updateZoomDisplay();
        }
        if (this.SN.Event.hasListeners('afterZoomChange')) {
            this.SN.Event.emit('afterZoomChange', { sheet: this, oldZoom, newZoom: next });
        }
    }

    /** @param {number} [step=0.1] @returns {number} */
    zoomIn(step = 0.1) {
        this.zoom = this.zoom + step;
        return this.zoom;
    }

    /** @param {number} [step=0.1] @returns {number} */
    zoomOut(step = 0.1) {
        this.zoom = this.zoom - step;
        return this.zoom;
    }

    /** @returns {number} Zoom to fit the selected area. */
    zoomToSelection() {
        const areas = this.activeAreas ?? [];
        let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
        if (areas.length === 0) {
            const active = this.activeCell;
            if (!active) return;
            minR = maxR = active.r;
            minC = maxC = active.c;
        } else {
            areas.forEach(area => {
                if (!area?.s || !area?.e) return;
                minR = Math.min(minR, area.s.r, area.e.r);
                minC = Math.min(minC, area.s.c, area.e.c);
                maxR = Math.max(maxR, area.s.r, area.e.r);
                maxC = Math.max(maxC, area.s.c, area.e.c);
            });
        }
        if (!Number.isFinite(minR) || !Number.isFinite(minC)) return;

        let totalWidth = 0;
        for (let c = minC; c <= maxC; c++) {
            const col = this.getCol(c);
            if (!col.hidden) totalWidth += col.width;
        }
        let totalHeight = 0;
        for (let r = minR; r <= maxR; r++) {
            const row = this.getRow(r);
            if (!row.hidden) totalHeight += row.height;
        }

        const viewWidth = Math.max(0, this.SN.Canvas.viewWidth - this.indexWidth);
        const viewHeight = Math.max(0, this.SN.Canvas.viewHeight - this.headHeight);
        if (totalWidth <= 0 || totalHeight <= 0 || viewWidth <= 0 || viewHeight <= 0) return;

        const targetZoom = Math.min(viewWidth / totalWidth, viewHeight / totalHeight);
        if (!Number.isFinite(targetZoom) || targetZoom <= 0) return;

        let rowStart = minR;
        let colStart = minC;
        if (this.getRow(minR)?.hidden) rowStart = this._findNextShowRow(minR - 1, 1);
        if (this.getCol(minC)?.hidden) colStart = this._findNextShowCol(minC - 1, 1);
        rowStart = Number.isFinite(rowStart) ? Math.max(0, rowStart) : 0;
        colStart = Number.isFinite(colStart) ? Math.max(0, colStart) : 0;

        if (!this.vi) this.vi = { rowArr: [rowStart], colArr: [colStart] };
        if (!Array.isArray(this.vi.rowArr)) this.vi.rowArr = [rowStart];
        if (!Array.isArray(this.vi.colArr)) this.vi.colArr = [colStart];
        this.vi.rowArr[0] = rowStart;
        this.vi.colArr[0] = colStart;
        this.vi.scrollTop = this.getScrollTop(rowStart);
        this.vi.scrollLeft = this.getScrollLeft(colStart);
        this.activeCell = { r: minR, c: minC };

        const prevZoom = this.zoom;
        this.zoom = targetZoom;
        if (this === this.SN.activeSheet && prevZoom === this.zoom) {
            this.SN._r();
        }
        return this.zoom;
    }

    /** @type {CellNum} */
    get viewStart() {
        return {
            r: this.vi?.rowArr?.[0] ?? 0,
            c: this.vi?.colArr?.[0] ?? 0
        };
    }

    set viewStart(val) {
        if (!this.vi) this.vi = { rowArr: [0], colArr: [0] };
        if (val.r !== undefined) {
            this.vi.rowArr = [val.r];
            this.vi.scrollTop = this.getScrollTop(val.r);
        }
        if (val.c !== undefined) {
            this.vi.colArr = [val.c];
            this.vi.scrollLeft = this.getScrollLeft(val.c);
        }
    }

    /** @type {number} px */
    get defaultColWidth() {
        return this._defaultColWidth;
    }

    set defaultColWidth(value) {
        const num = Number(value);
        const next = Number.isFinite(num) ? Math.max(1, Math.round(num)) : this._defaultColWidth;
        if (next === this._defaultColWidth) return;
        const prev = this._defaultColWidth;
        this._defaultColWidth = next;
        this._totalWidthCache = undefined;
        this._colAxisMetricsCache = null;
        this._colMetricVersion = (this._colMetricVersion || 0) + 1;
        this.cols.forEach(col => {
            if (col && col._width === prev) col._width = next;
        });
    }

    /** @type {number} px */
    get defaultRowHeight() {
        return this._defaultRowHeight;
    }

    set defaultRowHeight(value) {
        const num = Number(value);
        const next = Number.isFinite(num) ? Math.max(1, Math.round(num)) : this._defaultRowHeight;
        if (next === this._defaultRowHeight) return;
        const prev = this._defaultRowHeight;
        this._defaultRowHeight = next;
        this._totalHeightCache = undefined;
        this._rowAxisMetricsCache = null;
        this._rowMetricVersion = (this._rowMetricVersion || 0) + 1;
        this.rows.forEach(row => {
            if (row && row._height === prev) row._height = next;
        });
    }

    /** @type {number} px */
    get indexWidth() {
        if (!this.showRowColHeaders) return 0;
        return this._indexWidth;
    }

    set indexWidth(num) {
        this._indexWidth = num
    }

    /** @type {number} px */
    get headHeight() {
        if (!this.showRowColHeaders) return 0;
        return this._headHeight;
    }

    set headHeight(num) {
        this._headHeight = num
    }

    /** @type {number} */
    get frozenCols() {
        return this._frozenCols
    }

    set frozenCols(num) {
        const oldNum = this._frozenCols
        this.SN.UndoRedo.add({
            undo: () => {
                this._frozenCols = oldNum
            },
            redo: () => {
                this._frozenCols = num
            }
        });
        this._frozenCols = num
    }

    /** @type {number} */
    get frozenRows() {
        return this._frozenRows
    }

    set frozenRows(num) {
        const oldNum = this._frozenRows
        this.SN.UndoRedo.add({
            undo: () => {
                this._frozenRows = oldNum
            },
            redo: () => {
                this._frozenRows = num
            }
        });
        this._frozenRows = num
    }

    /** @type {boolean} */
    get hidden() {
        return this._hidden
    }

    set hidden(val) {
        const next = !!val;
        if (this.hidden == next) return
        if (next && this.SN.sheets.filter(st => !st.hidden).length == 1) return this.SN.Utils.toast(this.SN.t('core.sheet.sheet.toast.msg001'));
        const oldHidden = this._hidden;
        const eventBase = next ? 'SheetHide' : 'SheetShow';
        if (this.SN.Event.hasListeners(`before${eventBase}`)) {
            const beforeEvent = this.SN.Event.emit(`before${eventBase}`, {
                sheet: this,
                oldHidden,
                newHidden: next
            });
            if (beforeEvent.canceled) return;
        }
        this._hidden = next
        if (this.SN.activeSheet.name == this.name) this.SN.activeSheet = this.SN.sheets.find(s => !s.hidden);
        if (this.SN.Event.hasListeners(`after${eventBase}`)) {
            this.SN.Event.emit(`after${eventBase}`, {
                sheet: this,
                oldHidden,
                newHidden: next
            });
        }
    }

    /** @type {string} */
    get name() {
        return this._name
    }

    set name(newName) {
        const oldName = this._name;
        if (oldName === newName) return;
        if (!this.SN._validateSheetName(newName, this)) return;
        if (this.SN.Event.hasListeners('beforeSheetRename')) {
            const beforeEvent = this.SN.Event.emit('beforeSheetRename', { sheet: this, oldName, newName });
            if (beforeEvent.canceled) return;
        }
        this._name = newName
        this.SN._updListDom();
        if (this.SN.Event.hasListeners('afterSheetRename')) {
            this.SN.Event.emit('afterSheetRename', { sheet: this, oldName, newName });
        }
    }

    /** @type {number} */
    get rowCount() {
        return Math.max(this.rows.length, this._virtualRowCount || 0)
    }

    /** @type {number} */
    get colCount() {
        return Math.max(this.cols.length, this._virtualColCount || 0)
    }

    _setPendingXlsxRows(rows, meta = null) {
        const list = Array.isArray(rows) ? rows : [];
        const rowMap = new Map();
        const indexes = [];
        const stateMap = new Map();
        let maxRow = Math.max(-1, Number(meta?.rowCount) - 1 || -1);
        let maxCol = Math.max(-1, Number(meta?.colCount) - 1 || -1);
        const shouldScanCells = !Number.isFinite(Number(meta?.colCount));

        list.forEach((row, fallbackIndex) => {
            const rIndex = Math.max(0, (Number.parseInt(row?.['_$r'], 10) || (fallbackIndex + 1)) - 1);
            rowMap.set(rIndex, row);
            indexes.push(rIndex);
            maxRow = Math.max(maxRow, rIndex);

            if (shouldScanCells) {
                const cells = Array.isArray(row?.c) ? row.c : (row?.c ? [row.c] : []);
                cells.forEach(cell => {
                    const ref = cell?.['_$r'];
                    const match = typeof ref === 'string' ? ref.match(/[A-Z]+/i) : null;
                    if (!match) return;
                    maxCol = Math.max(maxCol, this.Utils.charToNum(match[0]));
                });
            }
        });

        (Array.isArray(meta?.customRows) ? meta.customRows : []).forEach(item => {
            const r = Math.max(0, Number(item?.r) || 0);
            const height = Number(item?.ht);
            const styleIndex = item?.styleIndex;
            stateMap.set(r, {
                height: Number.isFinite(height) && height > 0 ? Math.round(height * 1.333) : this.defaultRowHeight,
                hidden: !!item?.hidden,
                outlineLevel: Math.max(0, Math.min(7, Math.floor(Number(item?.outlineLevel) || 0))),
                collapsed: !!item?.collapsed,
                rowStyle: styleIndex !== undefined && item?.customFormat
                    ? this.SN.Xml._xGetStyle(styleIndex)
                    : null
            });
        });

        (Array.isArray(meta?.sharedFormulas) ? meta.sharedFormulas : []).forEach(item => {
            if (item?.si === undefined || !item.f) return;
            this.SN._sharedFormula[item.si] = {
                m: { r: item.r, c: item.c },
                f: item.f
            };
        });

        this._pendingXlsxRows = list;
        this._pendingXlsxRowMap = rowMap;
        this._pendingXlsxRowIndexes = indexes;
        this._pendingXlsxRowMeta = meta || null;
        this._pendingXlsxRowStateMap = stateMap.size > 0 ? stateMap : null;
        this._virtualRowCount = Math.max(this._virtualRowCount || 0, maxRow + 1, Number(meta?.rowCount) || 0);
        this._virtualColCount = Math.max(this._virtualColCount || 0, maxCol + 1, Number(meta?.colCount) || 0);
    }

    _getPendingXlsxRowState(r) {
        return this._pendingXlsxRowStateMap?.get(r) || null;
    }

    _ensurePendingXlsxRowState(r) {
        if (!this._pendingXlsxRowStateMap) this._pendingXlsxRowStateMap = new Map();
        let state = this._pendingXlsxRowStateMap.get(r);
        if (!state) {
            state = {
                height: this.defaultRowHeight,
                hidden: false,
                filterHidden: false,
                outlineLevel: 0,
                collapsed: false,
                rowStyle: null
            };
            this._pendingXlsxRowStateMap.set(r, state);
        }
        return state;
    }

    _applyPendingXlsxRowState(row, r) {
        const state = this._pendingXlsxRowStateMap?.get(r);
        if (!row || !state) return row;
        if (Number.isFinite(state.height) && state.height > 0) row._height = state.height;
        row._hidden = !!state.hidden;
        row._filterHidden = !!state.filterHidden;
        row._outlineLevel = state.outlineLevel || 0;
        row._collapsed = !!state.collapsed;
        if (state.rowStyle) row._rowStyle = state.rowStyle;
        this._pendingXlsxRowStateMap.delete(r);
        if (this._pendingXlsxRowStateMap.size === 0) this._pendingXlsxRowStateMap = null;
        return row;
    }

    _materializePendingRow(r) {
        const rowMap = this._pendingXlsxRowMap;
        if (!rowMap?.has(r)) return null;
        const rowXml = rowMap.get(r);
        const row = new Row(rowXml, this, r);
        this.rows[r] = this._applyPendingXlsxRowState(row, r);
        rowMap.delete(r);
        if (rowMap.size === 0) {
            this._pendingXlsxRows = null;
            this._pendingXlsxRowMap = null;
            this._pendingXlsxRowIndexes = null;
        }
        return this.rows[r];
    }

    _flushPendingXlsxRows() {
        if (!this._pendingXlsxRowMap || this._pendingXlsxRowMap.size === 0) return;
        const indexes = this._pendingXlsxRowIndexes?.length
            ? [...this._pendingXlsxRowIndexes]
            : [...this._pendingXlsxRowMap.keys()];
        indexes.sort((a, b) => a - b);
        indexes.forEach(r => {
            if (!this.rows[r]) this._materializePendingRow(r);
        });
    }

    async _flushPendingXlsxRowsAsync(options = {}) {
        if (!this._pendingXlsxRowMap || this._pendingXlsxRowMap.size === 0) return;
        const batchSize = Math.max(100, Number(options.batchSize) || 1500);
        const indexes = this._pendingXlsxRowIndexes?.length
            ? [...this._pendingXlsxRowIndexes]
            : [...this._pendingXlsxRowMap.keys()];
        indexes.sort((a, b) => a - b);

        for (let i = 0; i < indexes.length; i++) {
            const r = indexes[i];
            if (!this.rows[r]) this._materializePendingRow(r);
            if ((i + 1) % batchSize === 0) {
                options.onProgress?.({ current: i + 1, total: indexes.length, sheet: this });
                await new Promise(resolve => setTimeout(resolve, 0));
            }
        }
        options.onProgress?.({ current: indexes.length, total: indexes.length, sheet: this });
    }

    _extendVirtualRows(count = 200) {
        const current = this.rowCount;
        const next = Math.min(this._maxVirtualRowCount, current + Math.max(1, count));
        if (next <= current) return false;
        this._virtualRowCount = next;
        this._clearSizeCache();
        return true;
    }

    _extendVirtualCols(count = 32) {
        const current = this.colCount;
        const next = Math.min(this._maxVirtualColCount, current + Math.max(1, count));
        if (next <= current) return false;
        this._virtualColCount = next;
        this._clearSizeCache();
        return true;
    }

    // ============================= Public Methods =============================

    /** @param {RangeStr} range @returns {RangeNum} */
    rangeStrToNum(range) {
        const parts = range.replace(/['$]/g, "").split(':');
        let start, end, startColNum, endColNum, startRowNum, endRowNum;
        // 判断是行格式还是列格式，或者是正常的单元格区域
        if (/^\d+$/.test(parts[0])) { // 行格式 1:2
            startRowNum = parseInt(parts[0], 10) - 1;
            endRowNum = parseInt(parts[1], 10) - 1;
            startColNum = 0;
            endColNum = this.colCount - 1;
        } else if (/^[A-Z]+$/.test(parts[0])) { // 列格式 A:B
            startColNum = this.Utils.charToNum(parts[0]);
            endColNum = this.Utils.charToNum(parts[1]);
            startRowNum = 0;
            endRowNum = this.rowCount - 1;
        } else {
            start = parts[0].match(/([A-Z]+)(\d+)/);
            end = parts[1] ? parts[1].match(/([A-Z]+)(\d+)/) : start; // 单单元格时 end = start
            startColNum = this.Utils.charToNum(start[1]);
            startRowNum = parseInt(start[2], 10) - 1;
            endColNum = this.Utils.charToNum(end[1]);
            endRowNum = parseInt(end[2], 10) - 1;
        }
        // 保证开始小于结束
        return {
            s: { c: Math.min(startColNum, endColNum), r: Math.min(startRowNum, endRowNum) },
            e: { c: Math.max(startColNum, endColNum), r: Math.max(startRowNum, endRowNum) }
        };
    }

    /** @param {number} r @returns {Row} */
    getRow(r) {
        if (!this.rows[r]) this._materializePendingRow(r);
        if (!this.rows[r]) this.rows[r] = new Row(null, this, r);
        this._applyPendingXlsxRowState(this.rows[r], r);
        this.rows[r].rIndex = r
        return this.rows[r];
    }

    /** @param {number} c @returns {Col} */
    getCol(c) {
        if (!this.cols[c]) this.cols[c] = new Col(null, this, c);
        return this.cols[c]
    }

    /** @param {number|string} r @param {number} [c] @returns {Cell} */
    getCell(r, c) {
        if (typeof r === 'string') {
            ({ r, c } = this.Utils.cellStrToNum(r));
        }
        const row = this.getRow(r);
        if (!row.cells[c]) {
            row.cells[c] = new Cell(null, row, c);
        }
        return row.cells[c];
    }

    /** @param {RangeRef|RangeRef[]} ranges @param {(r:number,c:number,area:RangeNum)=>void} callback @param {{reverse?:boolean,sparse?:boolean}} [options={}] */
    eachCells(ranges, callback, options = {}) {
        const { reverse = false, sparse = false } = options;
        if (!Array.isArray(ranges)) ranges = [ranges];
        const areas = ranges.map(r => typeof r === 'string' ? this.rangeStrToNum(r) : r);
        const visited = new Set();

        if (sparse) {
            this.rows.forEach((row, r) => {
                if (!row) return;
                row.cells.forEach((cell, c) => {
                    if (!cell) return;
                    const area = areas.find(({ s, e }) => r >= s.r && r <= e.r && c >= s.c && c <= e.c);
                    if (!area) return;
                    const key = (r << 16) | c;
                    if (visited.has(key)) return;
                    visited.add(key);
                    callback(r, c, area);
                });
            });
        } else {
            areas.forEach(({ s, e }) => {
                const [rStart, rEnd, rStep] = reverse ? [e.r, s.r, -1] : [s.r, e.r, 1];
                const [cStart, cEnd, cStep] = reverse ? [e.c, s.c, -1] : [s.c, e.c, 1];
                for (let r = rStart; reverse ? r >= rEnd : r <= rEnd; r += rStep) {
                    for (let c = cStart; reverse ? c >= cEnd : c <= cEnd; c += cStep) {
                        const key = (r << 16) | c;
                        if (visited.has(key)) continue;
                        visited.add(key);
                        callback(r, c, { s, e });
                    }
                }
            });
        }
    }

    /** @param {boolean} [keep=false] */
    setBrush(keep = false) {
        if (this._brush.area) {
            this._brush.area = null;
            this._brush.keep = false;
            this.SN.Layout?.refreshActiveButtons?.();
            return;
        }
        this._brush.keep = keep;
        this._brush.area = { ...this.activeAreas[this.activeAreas.length - 1] };
        // 缓存源区域格式数据（用于平铺复制）
        this._cacheSourceFormats();
        this.SN.Layout?.refreshActiveButtons?.();
    }

    /** @returns {boolean} */
    cancelBrush() {
        if (!this._brush.area) return false;
        this._brush.area = null;
        this._brush.keep = false;
        this._brush.formats = null;
        this.SN.Layout?.refreshActiveButtons?.();
        return true;
    }

    /** @param {RangeNum} targetArea @returns {boolean} */
    applyBrush(targetArea) {
        if (!this._brush.area || !this._brush.formats) return false;
        const { sourceRows, sourceCols, formats } = this._brush;

        // 遍历目标区域，使用取模实现平铺
        for (let r = targetArea.s.r; r <= targetArea.e.r; r++) {
            for (let c = targetArea.s.c; c <= targetArea.e.c; c++) {
                // 计算源格式的索引（取模实现平铺）
                const srcRow = (r - targetArea.s.r) % sourceRows;
                const srcCol = (c - targetArea.s.c) % sourceCols;
                const srcFmt = formats[srcRow]?.[srcCol];
                if (!srcFmt) continue;

                const targetCell = this.getCell(r, c);
                // 复制完整格式（保留值不变）
                if (srcFmt.style) targetCell.style = JSON.parse(JSON.stringify(srcFmt.style));
                if (srcFmt.numFmt) targetCell.numFmt = srcFmt.numFmt;
                if (srcFmt.border) targetCell.border = JSON.parse(JSON.stringify(srcFmt.border));
            }
        }

        // 非连续模式，清除格式刷
        if (!this._brush.keep) {
            this._brush.area = null;
            this._brush.formats = null;
            this.SN.Layout?.refreshActiveButtons?.();
        }
        return true;
    }

    /** @param {Object|string} moveArea @param {Object|string} targetArea @returns {boolean} */
    moveArea(moveArea, targetArea) {
        if (typeof moveArea == 'string') moveArea = this.rangeStrToNum(moveArea);
        if (typeof targetArea == 'string') targetArea = this.rangeStrToNum(targetArea);
        if (this.SN.Event.hasListeners('beforeMoveArea')) {
            const beforeEvent = this.SN.Event.emit('beforeMoveArea', { sheet: this, moveArea, targetArea });
            if (beforeEvent.canceled) return false;
        }
        const offR = targetArea.s.r - moveArea.s.r  // 行偏移值
        const offC = targetArea.s.c - moveArea.s.c  // 列偏移值

        const mer = []
        const copys = []
        this.eachCells(moveArea, (r, c) => {
            const cell = this.getCell(r, c); // 要移动的区域
            copys.push({ style: cell.style, editVal: cell.editVal, hyperlink: cell.hyperlink, dataValidation: cell.dataValidation });
            cell.style = {};
            cell.editVal = ""
            cell.hyperlink = null;
            cell.dataValidation = null;
            // 存入合并并解除
            if (cell.master && cell.master.r == r && cell.master.c == c) {
                mer.push(this.merges.find(m => m.s.r == cell.master.r && m.s.c == cell.master.c));
                this.unMergeCells({ r, c });
            }
        })

        // 将目标区域合并的单元格全部取消，并赋copy
        let index = 0;
        this.eachCells(targetArea, (r, c) => {
            const cell = this.getCell(r, c);
            if (cell.isMerged) this.unMergeCells({ r, c });
            cell.style = copys[index]?.style ?? {}
            cell.editVal = copys[index]?.editVal ?? null
            cell.hyperlink = copys[index]?.hyperlink ?? null
            cell.dataValidation = copys[index]?.dataValidation ?? null
            index++;
        })

        // 合并新区域
        mer.forEach(m => {
            m.s.r += offR
            m.s.c += offC
            m.e.r += offR
            m.e.c += offC
            this.mergeCells(m)
        })

        if (this.SN.Event.hasListeners('afterMoveArea')) {
            this.SN.Event.emit('afterMoveArea', { sheet: this, moveArea, targetArea });
        }
        return true;
    }

    /** @param {Object|string} [oArea] @param {Object|string} [targetArea] @param {string} [type='order'] Fill area with series/copy/format. */
    paddingArea(
        oArea = this.Canvas.lastPadding.oArea,
        targetArea = this.Canvas.lastPadding.targetArea,
        type = 'order'
    ) {
        if (typeof oArea == 'string') oArea = this.rangeStrToNum(oArea);
        if (typeof targetArea == 'string') targetArea = this.rangeStrToNum(targetArea);

        // 收集源区数据
        const copys = [];
        this.eachCells(oArea, (r, c) => {
            const cell = this.getCell(r, c);
            copys.push({
                type: cell.type,
                style: { ...cell.style }, // 拷贝避免引用污染
                editVal: cell.editVal,
                _editVal: cell._editVal,
                cellIndex: { r, c }
            })
        });

        // 确定方向
        const direction = targetArea.s.r >= oArea.s.r ? 'add' : 'reduce';
        const step = direction === 'add' ? 1 : -1;

        // 序列递增辅助函数
        const applyOrderValue = (copy, prevValue) => {
            let result;

            // 优先处理日期/时间类型
            if (["date", "dateTime"].includes(copy.type)) {
                // 如果有上一步结果，则继续在其基础上加减，否则从初始值开始
                const baseDate = prevValue ? new Date(prevValue) : new Date(copy._editVal);
                baseDate.setDate(baseDate.getDate() + step);
                result = baseDate.toLocaleString();
            } else if (copy.type === "time") {
                const baseDate = prevValue ? new Date(prevValue) : new Date(copy._editVal);
                baseDate.setHours(baseDate.getHours() + step);
                result = baseDate.toLocaleString();
            }
            // 数字类型
            else if (typeof copy.editVal === 'number') {
                const baseNum = typeof prevValue === 'number' ? prevValue : copy.editVal;
                result = baseNum + step;
            }
            // 带数字的字符串
            else if (typeof copy.editVal === 'string' && /\d+/.test(copy.editVal)) {
                const prevNumMatch = typeof prevValue === 'string' && prevValue.match(/(\d+)/);
                const prevNum = prevNumMatch ? parseInt(prevNumMatch[0], 10) : parseInt(copy.editVal.match(/(\d+)/)[0], 10);
                const newNum = prevNum + step;
                result = copy.editVal.replace(/(\d+)/, newNum);
            }
            // 其他字符串保持不变
            else {
                result = copy.editVal;
            }

            return result;
        };

        // 遍历目标区域
        let i = 0;
        let prevValue = null; // 连续序列时基于上一步
        this.eachCells(targetArea, (r, c) => {
            const cell = this.getCell(r, c);
            if (cell.isMerged) this.unMergeCells({ r, c });

            const copy = copys[i];
            switch (type) {
                case 'order': {
                    const newVal = applyOrderValue(copy, prevValue);
                    cell.editVal = newVal;
                    cell.style = { ...copy.style };
                    prevValue = newVal; // 更新上一个值
                    break;
                }
                case 'copy': {
                    cell.style = { ...copy.style };
                    cell.editVal = copy.editVal;
                    break;
                }
                case 'format': {
                    cell.style = { ...copy.style };
                    cell.editVal = null;
                    break;
                }
                case 'noFormat': {
                    cell.editVal = copy.editVal;
                    cell.style = {};
                    break;
                }
            }

            if (++i >= copys.length) i = 0; // 循环使用源区
        });
    }

    showAllHidRows() {
        this.rows.forEach(row => { if (row) row.hidden = false; });
    }

    showAllHidCols() {
        this.cols.forEach(col => { if (col) col.hidden = false; });
    }

    /** @param {string} position @param {Object|null} options @param {Array} [area] */
    areasBorder(position, options, area = this.activeAreas) {
        if (!Array.isArray(area)) area = [area];
        const hasValidStyle = options && typeof options === 'object' && options.style !== '';
        // Traverse each target cell in the range.
        this.eachCells(area, (r, c, range) => {
            const { s, e } = range;
            const borderObj = {};
            if (position === 'none') {
                borderObj.top = null;
                borderObj.left = null;
                borderObj.bottom = null;
                borderObj.right = null;
                borderObj.diagonal = null;
                borderObj.diagonalUp = null;
                borderObj.diagonalDown = null;
            } else if (position === 'out') {
                if (r === s.r) borderObj.top = options;
                if (c === s.c) borderObj.left = options;
                if (r === e.r) borderObj.bottom = options;
                if (c === e.c) borderObj.right = options;
            } else if (position === 'inner') {
                if (r > s.r) borderObj.top = options;
                if (c > s.c) borderObj.left = options;
            } else if (position === 'top' && r === s.r) {
                borderObj.top = options;
            } else if (position === 'left' && c === s.c) {
                borderObj.left = options;
            } else if (position === 'bottom' && r === e.r) {
                borderObj.bottom = options;
            } else if (position === 'right' && c === e.c) {
                borderObj.right = options;
            } else if (position === 'diagonalDown') {
                borderObj.diagonal = hasValidStyle ? options : null;
                borderObj.diagonalDown = hasValidStyle ? true : null;
            } else if (position === 'diagonalUp') {
                borderObj.diagonal = hasValidStyle ? options : null;
                borderObj.diagonalUp = hasValidStyle ? true : null;
            } else if (position === 'all') {
                borderObj.top = options;
                borderObj.left = options;
                borderObj.bottom = options;
                borderObj.right = options;
            } else {
                return;
            }
            if (!Object.keys(borderObj).length) return;
            this.getCell(r, c).border = borderObj;
            // Keep adjacent shared edges consistent when writing outer sides.
            if (r == s.r && r != 0 && borderObj.top) {
                // top
                this.getCell(r - 1, c).border = { bottom: null };
            } else if (r == e.r && r != this.rowCount - 1 && borderObj.bottom) {
                // bottom
                this.getCell(r + 1, c).border = { top: null };
            }
            if (c == s.c && c != 0 && borderObj.left) {
                // left
                this.getCell(r, c - 1).border = { right: null };
            } else if (c == e.c && c != this.colCount - 1 && borderObj.right) {
                // right
                this.getCell(r, c + 1).border = { left: null };
            }
        });

    }

    /** Jump to hyperlink target of active cell. */
    hyperlinkJump() {

        const cell = this.getCell(this.activeCell.r, this.activeCell.c)
        if (!cell.hyperlink) return this.SN.Canvas.hyperlinkJumpBtn.style.display = 'none';  // 隐藏元素
        const location = cell.hyperlink.location
        if (location) { // 跳转本文档指定位置
            let sheetObj = this
            let cellRef = null
            if (location.includes("!")) {
                let [name, cellAd] = location.split("!")
                if (name.includes(" ") && name.startsWith("'") && name.endsWith("'")) name = name.slice(1, -1);
                sheetObj = this.SN.getSheet(name)
                cellRef = cellAd
            } else {
                cellRef = location
            }
            const posCell = sheetObj.getCell(cellRef)
            sheetObj.activeCell = { r: posCell.row.rIndex, c: posCell.cIndex }
            sheetObj.activeAreas = []
            sheetObj.viewStart = { r: sheetObj.activeCell.r, c: sheetObj.activeCell.c }
            this.SN.activeSheet = sheetObj
        } else { // 根据rId获取target
            const target = cell?.hyperlink?.target
            window.open(target ?? '')
        }
        this.SN.Canvas.hyperlinkJumpBtn.style.display = 'none';  // 隐藏元素

    }

    _cacheSourceFormats() {
        if (!this._brush.area) return;
        const { s, e } = this._brush.area;
        const formats = [];
        for (let r = s.r; r <= e.r; r++) {
            const row = [];
            for (let c = s.c; c <= e.c; c++) {
                const cell = this.getCell(r, c);
                row.push({
                    style: cell.style ? JSON.parse(JSON.stringify(cell.style)) : {},
                    numFmt: cell.numFmt || null,
                    border: cell.border ? JSON.parse(JSON.stringify(cell.border)) : null
                });
            }
            formats.push(row);
        }
        this._brush.formats = formats;
        this._brush.sourceRows = e.r - s.r + 1;
        this._brush.sourceCols = e.c - s.c + 1;
    }
}

// ============================= Prototype Mixin =============================

Object.assign(Sheet.prototype, {

    rangeSort: Sort.rangeSort,
    flashFill: FlashFill.flashFill,
    textToColumns: TextToColumns.textToColumns,
    groupRows: Outline.groupRows,
    ungroupRows: Outline.ungroupRows,
    groupCols: Outline.groupCols,
    ungroupCols: Outline.ungroupCols,
    subtotal: Subtotal.subtotal,
    consolidate: Consolidate.consolidate,
    setDataValidation: DataValidation.setDataValidation,
    clearDataValidation: DataValidation.clearDataValidation,
    /** @param {RangeRef|RangeRef[]} range @param {string|Object} [ops] */
    clearRange: ClearRange.clearRange,

    getUsedRange: UsedRange.getUsedRange,

    /** @param {(ICellConfig|string|number)[][]} arr @param {CellRef} pos @param {{align?:string,border?:boolean,width?:number,height?:number}} [ops] @returns {RangeNum} */
    insertTemplate: InsertTable.insertTemplate,

    /**
     * @deprecated Use insertTemplate instead. insertTable has ambiguous naming and will be deprecated after 2026-05-27.
     * @param {(ICellConfig|string|number)[][]} arr
     * @param {CellRef} pos
     * @param {{align?:string,border?:boolean,width?:number,height?:number}} [ops]
     * @returns {RangeNum}
     */
    insertTable: InsertTable.insertTable,

    /** @param {number} startRow @param {number} endRow @returns {number} */
    autoFitRows: AutoFit.autoFitRows,

    /** @param {number} startCol @param {number} endCol @returns {number} */
    autoFitCols: AutoFit.autoFitCols,

    _getExpandMaxArea: Merge._getExpandMaxArea,

    /** @param {Object} area @returns {boolean} */
    areaHaveMerge: Merge.areaHaveMerge,

    /** @param {CellRef} cell */
    unMergeCells: Merge.unMergeCells,

    /** @param {RangeRef} range */
    mergeCells: Merge.mergeCells,

    /** @param {number} r @param {number} number */
    addRows: RowColHandler.addRows,

    /** @param {number} c @param {number} number */
    addCols: RowColHandler.addCols,

    /** @param {number} r @param {number} number */
    delRows: RowColHandler.delRows,

    /** @param {number} c @param {number} number */
    delCols: RowColHandler.delCols,

    /** @param {number} r @param {number} c @returns {Object} */
    getCellInViewInfo: ViewCalculator.getCellInViewInfo,

    /** @param {Object} area @returns {Object} */
    getAreaInviewInfo: ViewCalculator.getAreaInviewInfo,

    /** @returns {number} */
    getTotalHeight: ViewCalculator.getTotalHeight,

    /** @returns {number} */
    getTotalWidth: ViewCalculator.getTotalWidth,

    /** @param {number} rowIndex @returns {number} */
    getScrollTop: ViewCalculator.getScrollTop,

    /** @param {number} colIndex @returns {number} */
    getScrollLeft: ViewCalculator.getScrollLeft,

    /** @param {number} scrollTop @returns {number} */
    getRowIndexByScrollTop: ViewCalculator.getRowIndexByScrollTop,

    /** @param {number} scrollLeft @returns {number} */
    getColIndexByScrollLeft: ViewCalculator.getColIndexByScrollLeft,

    _getAreaInviewInfo2: ViewCalculator._getAreaInviewInfo2,
    _updView: ViewCalculator._updView,
    _setViewScroll: ViewCalculator._setViewScroll,
    _findBeforeShowRow: ViewCalculator._findBeforeShowRow,
    _findBeforeShowCol: ViewCalculator._findBeforeShowCol,
    _findNextShowRow: ViewCalculator._findNextShowRow,
    _findNextShowCol: ViewCalculator._findNextShowCol,
    _clearSizeCache: ViewCalculator._clearSizeCache,

    // -------- 初始化 --------
    _init: Initializer._init,

    // -------- 剪贴板 --------
    /** @param {Array} [areas] */
    copy: Clipboard.copy,

    /** @param {Array} [areas] */
    cut: Clipboard.cut,

    /** @param {Object} [options] */
    paste: Clipboard.paste,

    clearClipboard: Clipboard.clearClipboard,

    /** @returns {Object} */
    getClipboardData: Clipboard.getClipboardData,

    /** @returns {boolean} */
    hasClipboardData: Clipboard.hasClipboardData,

    /** @param {string|Object|Array<string|Object>} range @param {Object} control */
    setCellControl: CellControl.setCellControl,

    /** @param {string|Object|Array<string|Object>} range */
    clearCellControl: CellControl.clearCellControl,

    /** @param {number} row @param {number} col @returns {boolean} */
    toggleCheckbox: CellControl.toggleCheckbox,

    /** @param {string|Object|Array<string|Object>} range @param {boolean} checked */
    setCheckboxValue: CellControl.setCheckboxValue,

    /** @param {string|Object|Array<string|Object>} range */
    deleteCheckboxControlOrValue: CellControl.deleteCheckboxControlOrValue,

    /** @param {string|Object|Array<string|Object>} range @returns {boolean} */
    hasCheckboxControl: CellControl.hasCheckboxControl
});
