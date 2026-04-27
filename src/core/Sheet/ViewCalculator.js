/**
 * View Calculating Module
 * Maintains the pixel-level viewport used by canvas rendering.
 */

const DEFAULT_ROW_EXTEND = 240;
const DEFAULT_COL_EXTEND = 48;
const EDGE_EXTEND_RATIO = 0.75;

function _clamp(value, min, max) {
    if (!Number.isFinite(value)) return min;
    return Math.min(Math.max(value, min), Math.max(min, max));
}

function _rowSize(sheet, index) {
    const row = sheet.rows[index];
    if (!row) return sheet.defaultRowHeight;
    return row.hidden ? 0 : row.height;
}

function _colSize(sheet, index) {
    const col = sheet.cols[index];
    if (!col) return sheet.defaultColWidth;
    return col.hidden ? 0 : col.width;
}

function _isRowVisible(sheet, index) {
    const row = sheet.rows[index];
    return !row || !row.hidden;
}

function _isColVisible(sheet, index) {
    const col = sheet.cols[index];
    return !col || !col.hidden;
}

function _findFirstVisible(count, isVisible) {
    for (let i = 0; i < count; i++) {
        if (isVisible(i)) return i;
    }
    return 0;
}

function _findLastVisible(count, isVisible) {
    for (let i = count - 1; i >= 0; i--) {
        if (isVisible(i)) return i;
    }
    return Math.max(0, count - 1);
}

function _sumRows(sheet, start, end) {
    let total = 0;
    for (let i = Math.max(0, start); i <= end && i < sheet.rowCount; i++) {
        total += _rowSize(sheet, i);
    }
    return total;
}

function _sumCols(sheet, start, end) {
    let total = 0;
    for (let i = Math.max(0, start); i <= end && i < sheet.colCount; i++) {
        total += _colSize(sheet, i);
    }
    return total;
}

function _getAxisPosition(sheet, index, axis) {
    const vi = sheet.vi || {};
    const isRow = axis === 'row';
    const frozenIndexes = isRow ? (vi.frozenRows || []) : (vi.frozenCols || []);
    const minFrozen = frozenIndexes[0];
    const maxFrozen = frozenIndexes[frozenIndexes.length - 1];
    const headerSize = isRow ? sheet.headHeight : sheet.indexWidth;
    const freezeSize = isRow ? (vi.fh ?? sheet.headHeight) : (vi.fw ?? sheet.indexWidth);
    const scrollPos = isRow ? (vi.scrollTop ?? 0) : (vi.scrollLeft ?? 0);
    const getOffset = isRow ? getScrollTop : getScrollLeft;

    if (frozenIndexes.length && index >= minFrozen && index <= maxFrozen) {
        return headerSize + (getOffset.call(sheet, index) - getOffset.call(sheet, minFrozen));
    }

    return freezeSize + (getOffset.call(sheet, index) - scrollPos);
}

function _hasLayoutInRange(layouts, startIndex, endIndex) {
    return layouts.some(item => item.index >= startIndex && item.index <= endIndex);
}

function _getVisibleBounds(layouts, startIndex, endIndex, posKey, sizeKey) {
    let start = Infinity;
    let end = -Infinity;
    for (const item of layouts) {
        if (item.index < startIndex || item.index > endIndex) continue;
        start = Math.min(start, item[posKey]);
        end = Math.max(end, item[posKey] + item[sizeKey]);
    }
    if (!Number.isFinite(start) || !Number.isFinite(end)) return null;
    return { start, end };
}

function _syncPixelScrollFromIndex(sheet, vi) {
    if (vi._pixelScrollDirty) {
        vi._pixelScrollDirty = false;
        return;
    }
    if (!Number.isFinite(vi.scrollTop)) {
        vi.scrollTop = getScrollTop.call(sheet, vi.rowArr?.[0] ?? 0);
    } else if (Number.isFinite(vi.rowArr?.[0]) && getRowIndexByScrollTop.call(sheet, vi.scrollTop) !== vi.rowArr[0]) {
        vi.scrollTop = getScrollTop.call(sheet, vi.rowArr[0]);
    }
    if (!Number.isFinite(vi.scrollLeft)) {
        vi.scrollLeft = getScrollLeft.call(sheet, vi.colArr?.[0] ?? 0);
    } else if (Number.isFinite(vi.colArr?.[0]) && getColIndexByScrollLeft.call(sheet, vi.scrollLeft) !== vi.colArr[0]) {
        vi.scrollLeft = getScrollLeft.call(sheet, vi.colArr[0]);
    }
}

function _getScrollBodySizes(sheet, viewWidth, viewHeight) {
    const frozenHeight = _sumRows(sheet, sheet._freezeStartRow || 0, (sheet._frozenRows || 0) - 1);
    const frozenWidth = _sumCols(sheet, sheet._freezeStartCol || 0, (sheet._frozenCols || 0) - 1);
    return {
        frozenHeight,
        frozenWidth,
        bodyHeight: Math.max(1, viewHeight - sheet.headHeight - frozenHeight),
        bodyWidth: Math.max(1, viewWidth - sheet.indexWidth - frozenWidth)
    };
}

function _getScrollBounds(sheet, viewWidth, viewHeight) {
    const { bodyHeight, bodyWidth } = _getScrollBodySizes(sheet, viewWidth, viewHeight);
    const minTop = getScrollTop.call(sheet, sheet._frozenRows || 0);
    const minLeft = getScrollLeft.call(sheet, sheet._frozenCols || 0);
    const maxTop = Math.max(minTop, getTotalHeight.call(sheet) - bodyHeight);
    const maxLeft = Math.max(minLeft, getTotalWidth.call(sheet) - bodyWidth);
    return { minTop, minLeft, maxTop, maxLeft, bodyHeight, bodyWidth };
}

function _maybeExtendVirtualSize(sheet, scrollLeft, scrollTop, viewWidth, viewHeight) {
    let bounds = _getScrollBounds(sheet, viewWidth, viewHeight);
    for (let i = 0; i < 8 && scrollTop > bounds.maxTop - bounds.bodyHeight * EDGE_EXTEND_RATIO; i++) {
        if (!sheet._extendVirtualRows?.(DEFAULT_ROW_EXTEND)) break;
        bounds = _getScrollBounds(sheet, viewWidth, viewHeight);
    }
    for (let i = 0; i < 8 && scrollLeft > bounds.maxLeft - bounds.bodyWidth * EDGE_EXTEND_RATIO; i++) {
        if (!sheet._extendVirtualCols?.(DEFAULT_COL_EXTEND)) break;
        bounds = _getScrollBounds(sheet, viewWidth, viewHeight);
    }
    return bounds;
}

function _buildRowLayouts(sheet, startIndex, startY, viewHeight, overscan) {
    const layouts = [];
    const indexes = [];
    let y = startY;
    for (let i = startIndex; i < sheet.rowCount; i++) {
        const h = _rowSize(sheet, i);
        if (h <= 0) continue;
        layouts.push({ index: i, y, h, frozen: false });
        indexes.push(i);
        y += h;
        if (y > viewHeight + overscan) break;
    }
    return { layouts, indexes, end: y };
}

function _buildColLayouts(sheet, startIndex, startX, viewWidth, overscan) {
    const layouts = [];
    const indexes = [];
    let x = startX;
    for (let i = startIndex; i < sheet.colCount; i++) {
        const w = _colSize(sheet, i);
        if (w <= 0) continue;
        layouts.push({ index: i, x, w, frozen: false });
        indexes.push(i);
        x += w;
        if (x > viewWidth + overscan) break;
    }
    return { layouts, indexes, end: x };
}

/**
 * Fetch cell information in visible view.
 * @param {Number} rowIndex - Row index
 * @param {Number} colIndex - Column index
 * @param {Boolean} posMerge - Whether to process merging cells
 * @returns {Object}
 */
export function getCellInViewInfo(rowIndex, colIndex, posMerge = true) {
    let r = rowIndex;
    let c = colIndex;
    let merge;

    if (posMerge) {
        const cell = this.getCell(rowIndex, colIndex);
        if (cell.master) {
            r = cell.master.r;
            c = cell.master.c;
            merge = this.merges.find(m => m.s.r === r && m.s.c === c);
        }
    }

    const vi = this.vi || {};
    const rowLayout = vi.rowMap?.get(r);
    const colLayout = vi.colMap?.get(c);
    const x = colLayout?.x ?? _getAxisPosition(this, c, 'col');
    const y = rowLayout?.y ?? _getAxisPosition(this, r, 'row');
    const w = merge ? _sumCols(this, merge.s.c, merge.e.c) : _colSize(this, colIndex);
    const h = merge ? _sumRows(this, merge.s.r, merge.e.r) : _rowSize(this, rowIndex);
    const viewWidth = this.Canvas?.viewWidth ?? this._handleLayer.offsetWidth;
    const viewHeight = this.Canvas?.viewHeight ?? this._handleLayer.offsetHeight;
    const inView = x + w > this.indexWidth && y + h > this.headHeight && x < viewWidth && y < viewHeight;

    return { inView, x, y, w, h };
}

/**
 * Fetch clipped area information in visible view.
 * @param {Object} area - Area object
 * @returns {Object}
 */
export function getAreaInviewInfo(area) {
    const startCol = Math.min(area.s.c, area.e.c);
    const endCol = Math.max(area.s.c, area.e.c);
    const startRow = Math.min(area.s.r, area.e.r);
    const endRow = Math.max(area.s.r, area.e.r);
    const vi = this.vi || {};
    const rowBounds = _getVisibleBounds(vi.rowLayouts || [], startRow, endRow, 'y', 'h');
    const colBounds = _getVisibleBounds(vi.colLayouts || [], startCol, endCol, 'x', 'w');
    if (!rowBounds || !colBounds) return { inView: false };

    return {
        x: colBounds.start,
        y: rowBounds.start,
        w: colBounds.end - colBounds.start,
        h: rowBounds.end - rowBounds.start,
        inView: true
    };
}

/**
 * Get full area information in visible view. Coordinates may be negative when
 * the area starts outside the viewport.
 * @param {Object} area - Area object
 * @returns {Object}
 */
export function _getAreaInviewInfo2(area) {
    const startCol = Math.min(area.s.c, area.e.c);
    const endCol = Math.max(area.s.c, area.e.c);
    const startRow = Math.min(area.s.r, area.e.r);
    const endRow = Math.max(area.s.r, area.e.r);
    const vi = this.vi || {};
    const rowInView = _hasLayoutInRange(vi.rowLayouts || [], startRow, endRow);
    const colInView = _hasLayoutInRange(vi.colLayouts || [], startCol, endCol);
    if (!rowInView || !colInView) return { inView: false };

    return {
        x: _getAxisPosition(this, startCol, 'col'),
        y: _getAxisPosition(this, startRow, 'row'),
        w: _sumCols(this, startCol, endCol),
        h: _sumRows(this, startRow, endRow),
        inView: true
    };
}

/**
 * Update visible indexes and pixel layouts.
 * @param {Number} canvasWidth - Canvas logical width
 * @param {Number} canvasHeight - Canvas logical height
 */
export function _updView(canvasWidth, canvasHeight) {
    const viewWidth = canvasWidth ?? this.Canvas?.viewWidth ?? this._handleLayer.offsetWidth;
    const viewHeight = canvasHeight ?? this.Canvas?.viewHeight ?? this._handleLayer.offsetHeight;
    const rowCount = Math.max(1, this.rowCount);
    const colCount = Math.max(1, this.colCount);
    const side = {
        t: _findFirstVisible(rowCount, index => _isRowVisible(this, index)),
        l: _findFirstVisible(colCount, index => _isColVisible(this, index)),
        b: _findLastVisible(rowCount, index => _isRowVisible(this, index)),
        r: _findLastVisible(colCount, index => _isColVisible(this, index))
    };

    const vi = this.vi ?? { rowArr: [side.t], colArr: [side.l] };
    if (!Array.isArray(vi.rowArr)) vi.rowArr = [side.t];
    if (!Array.isArray(vi.colArr)) vi.colArr = [side.l];
    _syncPixelScrollFromIndex(this, vi);

    const bounds = _maybeExtendVirtualSize(this, vi.scrollLeft, vi.scrollTop, viewWidth, viewHeight);
    vi.scrollTop = _clamp(vi.scrollTop, bounds.minTop, bounds.maxTop);
    vi.scrollLeft = _clamp(vi.scrollLeft, bounds.minLeft, bounds.maxLeft);

    const overscan = Math.max(48, this.Canvas?._renderOverscan ?? 96);
    const newVi = {
        frozenCols: [],
        frozenRows: [],
        rowArr: [],
        colArr: [],
        rowsIndex: [],
        colsIndex: [],
        rowLayouts: [],
        colLayouts: [],
        rowMap: new Map(),
        colMap: new Map(),
        offsetY: 0,
        offsetX: 0,
        scrollTop: vi.scrollTop,
        scrollLeft: vi.scrollLeft,
        w: this.indexWidth,
        h: this.headHeight,
        fw: this.indexWidth,
        fh: this.headHeight,
        side
    };

    const freezeStartCol = this._freezeStartCol || 0;
    const freezeStartRow = this._freezeStartRow || 0;
    const frozenCols = this._frozenCols || 0;
    const frozenRows = this._frozenRows || 0;

    let x = this.indexWidth;
    for (let c = freezeStartCol; c < Math.min(frozenCols, colCount); c++) {
        const w = _colSize(this, c);
        if (w <= 0) continue;
        const item = { index: c, x, w, frozen: true };
        newVi.frozenCols.push(c);
        newVi.colLayouts.push(item);
        newVi.colMap.set(c, item);
        x += w;
    }
    newVi.fw = x;

    let y = this.headHeight;
    for (let r = freezeStartRow; r < Math.min(frozenRows, rowCount); r++) {
        const h = _rowSize(this, r);
        if (h <= 0) continue;
        const item = { index: r, y, h, frozen: true };
        newVi.frozenRows.push(r);
        newVi.rowLayouts.push(item);
        newVi.rowMap.set(r, item);
        y += h;
    }
    newVi.fh = y;

    const scrollRowStart = Math.max(frozenRows, getRowIndexByScrollTop.call(this, vi.scrollTop));
    const scrollColStart = Math.max(frozenCols, getColIndexByScrollLeft.call(this, vi.scrollLeft));
    const rowStartTop = getScrollTop.call(this, scrollRowStart);
    const colStartLeft = getScrollLeft.call(this, scrollColStart);
    newVi.offsetY = vi.scrollTop - rowStartTop;
    newVi.offsetX = vi.scrollLeft - colStartLeft;

    const rowBuild = _buildRowLayouts(this, scrollRowStart, newVi.fh - newVi.offsetY, viewHeight, overscan);
    rowBuild.layouts.forEach(item => {
        newVi.rowArr.push(item.index);
        newVi.rowLayouts.push(item);
        newVi.rowMap.set(item.index, item);
    });

    const colBuild = _buildColLayouts(this, scrollColStart, newVi.fw - newVi.offsetX, viewWidth, overscan);
    colBuild.layouts.forEach(item => {
        newVi.colArr.push(item.index);
        newVi.colLayouts.push(item);
        newVi.colMap.set(item.index, item);
    });

    if (newVi.rowArr.length === 0) {
        const item = { index: scrollRowStart, y: newVi.fh, h: _rowSize(this, scrollRowStart), frozen: false };
        newVi.rowArr.push(item.index);
        newVi.rowLayouts.push(item);
        newVi.rowMap.set(item.index, item);
    }
    if (newVi.colArr.length === 0) {
        const item = { index: scrollColStart, x: newVi.fw, w: _colSize(this, scrollColStart), frozen: false };
        newVi.colArr.push(item.index);
        newVi.colLayouts.push(item);
        newVi.colMap.set(item.index, item);
    }

    newVi.rowsIndex = [...newVi.frozenRows, ...newVi.rowArr];
    newVi.colsIndex = [...newVi.frozenCols, ...newVi.colArr];
    newVi.h = Math.max(this.headHeight, ...newVi.rowLayouts.map(item => item.y + item.h));
    newVi.w = Math.max(this.indexWidth, ...newVi.colLayouts.map(item => item.x + item.w));
    this.vi = newVi;
}

/**
 * Set pixel scroll position and refresh viewport indexes.
 * @param {Number} scrollLeft - Horizontal scroll pixels
 * @param {Number} scrollTop - Vertical scroll pixels
 * @param {Number} canvasWidth - Canvas logical width
 * @param {Number} canvasHeight - Canvas logical height
 * @returns {boolean}
 */
export function _setViewScroll(scrollLeft, scrollTop, canvasWidth, canvasHeight) {
    if (!this.vi) this.vi = { rowArr: [0], colArr: [0] };
    _syncPixelScrollFromIndex(this, this.vi);
    const viewWidth = canvasWidth ?? this.Canvas?.viewWidth ?? this._handleLayer.offsetWidth;
    const viewHeight = canvasHeight ?? this.Canvas?.viewHeight ?? this._handleLayer.offsetHeight;
    const bounds = _maybeExtendVirtualSize(this, scrollLeft, scrollTop, viewWidth, viewHeight);
    const nextLeft = _clamp(scrollLeft, bounds.minLeft, bounds.maxLeft);
    const nextTop = _clamp(scrollTop, bounds.minTop, bounds.maxTop);
    const changed = Math.abs((this.vi.scrollLeft ?? 0) - nextLeft) > 0.01 ||
        Math.abs((this.vi.scrollTop ?? 0) - nextTop) > 0.01;
    this.vi.scrollLeft = nextLeft;
    this.vi.scrollTop = nextTop;
    this.vi._pixelScrollDirty = true;
    this._updView(viewWidth, viewHeight);
    return changed;
}

/**
 * Find index forward for rows not hidden.
 * @param {Number} r - Current row index
 * @param {Number} number - Distance to search forward
 * @returns {Number}
 */
export function _findBeforeShowRow(r, number) {
    let count = 0;
    let res = r;
    for (let i = Math.min(r - 1, this.rowCount - 1); i >= 0; i--) {
        if (_isRowVisible(this, i)) {
            count++;
            res = i;
        }
        if (count >= number) return i;
    }
    return Math.max(0, res);
}

/**
 * Find index forward for columns not hidden.
 * @param {Number} c - Current column index
 * @param {Number} number - Distance to search forward
 * @returns {Number}
 */
export function _findBeforeShowCol(c, number) {
    let count = 0;
    let res = c;
    for (let i = Math.min(c - 1, this.colCount - 1); i >= 0; i--) {
        if (_isColVisible(this, i)) {
            count++;
            res = i;
        }
        if (count >= number) return i;
    }
    return Math.max(0, res);
}

/**
 * Find index back to specified non-hidden rows.
 * @param {Number} r - Current row index
 * @param {Number} number - Distance to search back
 * @returns {Number}
 */
export function _findNextShowRow(r, number) {
    let count = 0;
    let res = r;
    for (let i = r + 1; i < this.rowCount; i++) {
        if (_isRowVisible(this, i)) {
            count++;
            res = i;
        }
        if (count >= number) return i;
    }
    this._extendVirtualRows?.(DEFAULT_ROW_EXTEND);
    return Math.min(this.rowCount - 1, res + Math.max(1, number - count));
}

/**
 * Find index back for columns not hidden.
 * @param {Number} c - Current column index
 * @param {Number} number - Distance to search back
 * @returns {Number}
 */
export function _findNextShowCol(c, number) {
    let count = 0;
    let res = c;
    for (let i = c + 1; i < this.colCount; i++) {
        if (_isColVisible(this, i)) {
            count++;
            res = i;
        }
        if (count >= number) return i;
    }
    this._extendVirtualCols?.(DEFAULT_COL_EXTEND);
    return Math.min(this.colCount - 1, res + Math.max(1, number - count));
}

/**
 * Retrieving the total height of all rows.
 * @returns {Number}
 */
export function getTotalHeight() {
    if (this._totalHeightCache !== undefined) return this._totalHeightCache;
    let total = this.rowCount * this.defaultRowHeight;
    this.rows.forEach((row, index) => {
        if (!row || index >= this.rowCount) return;
        total += (row.hidden ? 0 : row.height) - this.defaultRowHeight;
    });
    this._totalHeightCache = Math.max(0, total);
    return this._totalHeightCache;
}

/**
 * Get the total width of all columns.
 * @returns {Number}
 */
export function getTotalWidth() {
    if (this._totalWidthCache !== undefined) return this._totalWidthCache;
    let total = this.colCount * this.defaultColWidth;
    this.cols.forEach((col, index) => {
        if (!col || index >= this.colCount) return;
        total += (col.hidden ? 0 : col.width) - this.defaultColWidth;
    });
    this._totalWidthCache = Math.max(0, total);
    return this._totalWidthCache;
}

/**
 * Clear size cache.
 */
export function _clearSizeCache() {
    this._totalHeightCache = undefined;
    this._totalWidthCache = undefined;
}

/**
 * Get the sum of heights before the specified row.
 * @param {Number} rowIndex - Row index
 * @returns {Number}
 */
export function getScrollTop(rowIndex) {
    const limit = Math.max(0, Math.min(rowIndex, this.rowCount));
    let total = limit * this.defaultRowHeight;
    this.rows.forEach((row, index) => {
        if (!row || index >= limit) return;
        total += (row.hidden ? 0 : row.height) - this.defaultRowHeight;
    });
    return Math.max(0, total);
}

/**
 * Sum of widths before the specified column.
 * @param {Number} colIndex - Column index
 * @returns {Number}
 */
export function getScrollLeft(colIndex) {
    const limit = Math.max(0, Math.min(colIndex, this.colCount));
    let total = limit * this.defaultColWidth;
    this.cols.forEach((col, index) => {
        if (!col || index >= limit) return;
        total += (col.hidden ? 0 : col.width) - this.defaultColWidth;
    });
    return Math.max(0, total);
}

/**
 * Find row index from vertical scroll pixel position.
 * @param {Number} scrollTop - Scroll pixels position
 * @returns {Number}
 */
export function getRowIndexByScrollTop(scrollTop) {
    if (this.rowCount <= 0) return 0;
    const target = _clamp(scrollTop, 0, Math.max(0, getTotalHeight.call(this) - 1));
    let low = 0;
    let high = this.rowCount - 1;
    let result = high;
    while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        if (getScrollTop.call(this, mid + 1) > target) {
            result = mid;
            high = mid - 1;
        } else {
            low = mid + 1;
        }
    }
    while (result < this.rowCount - 1 && !_isRowVisible(this, result)) result++;
    return Math.max(0, Math.min(result, this.rowCount - 1));
}

/**
 * Find column index from horizontal scroll pixel position.
 * @param {Number} scrollLeft - Scroll pixels position
 * @returns {Number}
 */
export function getColIndexByScrollLeft(scrollLeft) {
    if (this.colCount <= 0) return 0;
    const target = _clamp(scrollLeft, 0, Math.max(0, getTotalWidth.call(this) - 1));
    let low = 0;
    let high = this.colCount - 1;
    let result = high;
    while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        if (getScrollLeft.call(this, mid + 1) > target) {
            result = mid;
            high = mid - 1;
        } else {
            low = mid + 1;
        }
    }
    while (result < this.colCount - 1 && !_isColVisible(this, result)) result++;
    return Math.max(0, Math.min(result, this.colCount - 1));
}
