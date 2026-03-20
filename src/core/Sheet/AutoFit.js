import { getCanvasFont, getFontSizePx } from '../Cell/fontDefaults.js';

const _MAX_AUTO_FIT_COL_WIDTH = 500;
const _AUTO_FIT_ROW_WRAP_PADDING = 8;
const _AUTO_FIT_ROW_EXTRA_HEIGHT = 6;
const _AUTO_FIT_COL_EXTRA_WIDTH = 16;

let _measureContext = null;

function _getMeasureContext() {
    if (_measureContext) return _measureContext;

    if (typeof OffscreenCanvas !== 'undefined') {
        const canvas = new OffscreenCanvas(1, 1);
        _measureContext = canvas.getContext('2d');
        if (_measureContext) return _measureContext;
    }

    if (typeof document !== 'undefined' && typeof document.createElement === 'function') {
        const canvas = document.createElement('canvas');
        _measureContext = canvas.getContext('2d');
    }

    return _measureContext;
}

function _normalizeIndexRange(start, end) {
    const startIndex = Math.max(0, Math.floor(Number(start) || 0));
    const endIndex = Math.max(0, Math.floor(Number(end) || 0));
    return startIndex <= endIndex
        ? { start: startIndex, end: endIndex }
        : { start: endIndex, end: startIndex };
}

function _hasCellContent(value) {
    return value !== undefined && value !== null && value !== '';
}

function _wrapText(text, maxWidth, context) {
    const lines = [];
    const words = String(text ?? '').split('');
    let currentLine = words[0] || '';

    for (let i = 1; i < words.length; i++) {
        const word = words[i];
        const width = context.measureText(currentLine + word).width;
        if (width <= maxWidth) {
            currentLine += word;
        } else {
            lines.push(currentLine);
            currentLine = word;
        }
    }
    lines.push(currentLine);

    return lines;
}

function _getFontInfo(font, source, fontCache) {
    const sizePx = getFontSizePx(font, source);
    const fontStr = getCanvasFont(font, source, sizePx);
    const key = fontStr;

    if (!fontCache.has(key)) fontCache.set(key, fontStr);

    return { key, str: fontStr, sizePx };
}

function _measureTextWidth(ctx, measureCache, fontKey, fontStr, text) {
    const normalizedText = String(text ?? '');
    const cacheKey = `${fontKey}|${normalizedText}`;
    if (measureCache.has(cacheKey)) return measureCache.get(cacheKey);

    ctx.font = fontStr;
    const width = ctx.measureText(normalizedText).width;
    measureCache.set(cacheKey, width);
    return width;
}

/**
 * Auto fit row heights in range.
 * @param {number} startRow
 * @param {number} endRow
 * @returns {number}
 */
export function autoFitRows(startRow, endRow) {
    const ctx = _getMeasureContext();
    if (!ctx) return 0;

    const { start, end } = _normalizeIndexRange(startRow, endRow);
    const fontCache = new Map();
    const measureCache = new Map();
    const colWidthCache = new Map();
    let changedCount = 0;

    for (let r = start; r <= end; r++) {
        const row = this.rows[r];
        if (!row || !row.cells?.length) continue;

        let maxHeight = this.defaultRowHeight;
        let hasContent = false;

        row.cells.forEach((cell, c) => {
            if (!cell) return;

            const value = cell.showVal;
            if (!_hasCellContent(value)) return;
            hasContent = true;

            const font = cell.font || {};
            const { key, str, sizePx } = _getFontInfo(font, this, fontCache);
            const colWidth = colWidthCache.has(c)
                ? colWidthCache.get(c)
                : (this.cols[c]?._width ?? this.defaultColWidth);
            colWidthCache.set(c, colWidth);

            const baseLines = String(value).split('\n');
            const lineCount = baseLines.reduce((count, line) => {
                if (cell.alignment?.wrapText) {
                    ctx.font = str;
                    return count + _wrapText(line, Math.max(1, colWidth - _AUTO_FIT_ROW_WRAP_PADDING), ctx).length;
                }
                return count + 1;
            }, 0);

            if (lineCount > 0) {
                const lineHeight = lineCount > 1 ? sizePx * 1.35 : sizePx * 1.15;
                const totalHeight = lineCount * lineHeight;
                maxHeight = Math.max(maxHeight, totalHeight + _AUTO_FIT_ROW_EXTRA_HEIGHT);
            }
        });

        if (!hasContent) continue;

        const newHeight = Math.round(maxHeight);
        const oldHeight = row.height;
        if (newHeight === oldHeight) continue;

        row.height = newHeight;
        if (row.height === newHeight) changedCount++;
    }

    return changedCount;
}

/**
 * Auto fit column widths in range.
 * @param {number} startCol
 * @param {number} endCol
 * @returns {number}
 */
export function autoFitCols(startCol, endCol) {
    const ctx = _getMeasureContext();
    if (!ctx) return 0;

    const { start, end } = _normalizeIndexRange(startCol, endCol);
    const fontCache = new Map();
    const measureCache = new Map();
    const maxWidthByCol = new Map();

    this.rows.forEach(row => {
        if (!row || !row.cells?.length) return;

        row.cells.forEach((cell, c) => {
            if (!cell || c < start || c > end) return;

            const value = cell.showVal;
            if (!_hasCellContent(value)) return;

            const font = cell.font || {};
            const { key, str } = _getFontInfo(font, this, fontCache);
            const lines = String(value).split('\n');
            let maxWidth = maxWidthByCol.get(c) || 0;

            for (const line of lines) {
                const width = _measureTextWidth(ctx, measureCache, key, str, line) + _AUTO_FIT_COL_EXTRA_WIDTH;
                if (width > maxWidth) maxWidth = width;
            }

            if (maxWidth > 0) maxWidthByCol.set(c, maxWidth);
        });
    });

    let changedCount = 0;

    maxWidthByCol.forEach((maxWidth, c) => {
        const col = this.getCol(c);
        const newWidth = Math.min(Math.round(maxWidth), _MAX_AUTO_FIT_COL_WIDTH);
        const oldWidth = col.width;
        if (newWidth === oldWidth) return;

        col.width = newWidth;
        if (col.width === newWidth) changedCount++;
    });

    return changedCount;
}
