/**
 * Input management module
 * Handles cell editor positioning, updates, and validation.
 */

import { getFormulaBarValue } from './helpers.js';
import { richTextToHTML, htmlToRichText, hasRichContent } from './RichTextEditor.js';

function _trValidationError(SN, text) {
    const raw = text || SN.t('dataValidation.errors.inputRuleMismatch');
    return SN.t(raw);
}

function _pickDefined(...values) {
    for (const value of values) {
        if (value !== undefined) return value;
    }
    return undefined;
}

function _getStripeColor(index, stripe1Color, stripe2Color, stripe1Size = 1, stripe2Size = 1) {
    const s1 = Math.max(1, stripe1Size || 1);
    const s2 = Math.max(1, stripe2Size || 1);
    const cycle = s1 + s2;
    const offset = index % cycle;
    return offset < s1 ? stripe1Color : stripe2Color;
}

function _getTableCellFillInfo(sheet, row, col) {
    if (!sheet?.Table?.size) return null;
    const table = sheet.Table.getTableAt(row, col);
    if (!table) return null;

    const style = table.style;
    if (!style) return null;

    const allowDxfOverlay = style._fromTokens !== true;
    const fromImport = table._fromImport === true;

    if (table.isHeaderCell(row, col)) {
        let bgColor = style.headerBg ?? null;
        if (allowDxfOverlay) {
            const dxfStyle = table._getHeaderDxfStyle?.();
            if (dxfStyle?.fill?.fgColor) {
                bgColor = dxfStyle.fill.fgColor;
            }
        }
        return { bgColor, fromImport };
    }

    if (table.isTotalsCell(row, col)) {
        return { bgColor: style.totalsBg ?? null, fromImport };
    }

    if (!table.isDataCell(row, col)) return null;

    const dataRowIndex = table.getDataRowIndex(row);
    const colIndex = table.getColumnIndex(col);
    let bgColor = style.dataBg ?? null;

    if (table.showRowStripes && dataRowIndex >= 0) {
        bgColor = _getStripeColor(
            dataRowIndex,
            style.rowStripe1,
            style.rowStripe2,
            style.rowStripe1Size,
            style.rowStripe2Size
        );
    }

    if (table.showColumnStripes && colIndex >= 0) {
        const stripe1Color = _pickDefined(style.columnStripe1, style.rowStripe2, style.rowStripe1, bgColor);
        const stripe2Color = _pickDefined(style.columnStripe2, style.rowStripe1, style.rowStripe2, bgColor);
        bgColor = _getStripeColor(
            colIndex,
            stripe1Color,
            stripe2Color,
            style.columnStripe1Size,
            style.columnStripe2Size
        );
    }

    if (table.showFirstColumn && colIndex === 0) {
        bgColor = _pickDefined(style.firstColumnBg, bgColor);
    }

    const lastColIndex = table.range.e.c - table.range.s.c;
    if (table.showLastColumn && colIndex === lastColIndex) {
        bgColor = _pickDefined(style.lastColumnBg, bgColor);
    }

    if (allowDxfOverlay) {
        const dxfStyle = table._getDataDxfStyle?.(colIndex);
        if (dxfStyle?.fill?.fgColor) {
            bgColor = dxfStyle.fill.fgColor;
        }
    }

    return { bgColor, fromImport };
}

function _getEditorBackgroundColor(sheet, cell, row, col, eyeProtectionMode, eyeProtectionColor) {
    const defaultBg = eyeProtectionMode ? eyeProtectionColor : '#ffffff';
    const cfFillColor = sheet.CF?.rules?.length ? sheet.CF.getFormat(row, col)?.fill?.fgColor : null;
    const tableFillInfo = _getTableCellFillInfo(sheet, row, col);
    const tableBgColor = tableFillInfo?.bgColor;
    const preferImportedTableStyle = tableFillInfo?.fromImport === true;
    const fill = cell.fill || {};

    if (cfFillColor) return cfFillColor;
    if (preferImportedTableStyle && tableBgColor) return tableBgColor;
    if (fill.type === 'pattern' && fill.pattern === 'solid' && fill.fgColor) return fill.fgColor;
    if (tableBgColor) return tableBgColor;

    if (fill.type === 'pattern') {
        if (fill.pattern === 'solid' || fill.pattern === 'none') {
            return fill.fgColor ?? defaultBg;
        }
        return fill.bgColor || defaultBg;
    }

    if (fill.type === 'gradient') {
        if (Array.isArray(fill.stops) && fill.stops.length > 0) {
            const firstStopWithColor = fill.stops.find(stop => stop?.color);
            if (firstStopWithColor?.color) return firstStopWithColor.color;
        }
        return fill.fgColor || fill.bgColor || defaultBg;
    }

    return fill.fgColor || defaultBg;
}

// Render floating cell editor
export function showCellInput(clearValue = false, focusEditor = false) {
    const sheet = this.activeSheet;
    const { c, r } = sheet.activeCell;
    const cell = sheet.getCell(r, c);
    const cInfo = this.activeSheet.getCellInViewInfo(r, c);
    const zoom = this.activeSheet.zoom ?? 1;
    const baseFontSize = (cell.font.size ?? 11) * 1.333;
    const cellW = cInfo.w * zoom;
    const cellH = cInfo.h * zoom;

    this.input.style.left = `${cInfo.x * zoom}px`;
    this.input.style.top = `${cInfo.y * zoom}px`;
    this.input.style.minWidth = `${Math.max(1, cellW - 1)}px`;
    this.input.style.minHeight = `${Math.max(1, cellH - 1)}px`;
    this.input.style.fontSize = `${baseFontSize * zoom}px`;
    this.input.style.fontFamily = cell.font.name ?? 'Calibri';

    // 水平对齐
    const hAlign = cell.horizontalAlign;
    this.input.style.textAlign = hAlign === 'center' ? 'center' : hAlign === 'right' ? 'right' : 'left';

    // 行高与 Canvas 渲染一致
    const lineHeightFactor = 1.35;
    this.input.style.lineHeight = `${baseFontSize * lineHeightFactor * zoom}px`;

    // 垂直对齐：通过 padding-top 模拟
    const vAlign = cell.verticalAlign;
    const textH = baseFontSize * lineHeightFactor * zoom;
    if (vAlign === 'center') {
        this.input.style.paddingTop = `${Math.max(0, (cellH - textH) / 2)}px`;
    } else if (vAlign === 'top') {
        this.input.style.paddingTop = `${2 * zoom}px`;
    } else {
        // bottom（默认）
        this.input.style.paddingTop = `${Math.max(0, cellH - textH - 1)}px`;
    }

    // 缩进
    const indent = cell.alignment?.indent || 0;
    this.input.style.paddingLeft = indent ? `${indent * 9 * zoom + 2}px` : '2px';
    this.input.style.paddingRight = '2px';

    // 单元格级别字体颜色
    this.input.style.color = cell.font.color || '';
    this.input.style.backgroundColor = _getEditorBackgroundColor(
        sheet,
        cell,
        r,
        c,
        this.eyeProtectionMode,
        this.eyeProtectionColor
    );

    this.inputEditing = !!cInfo.inView;
    if (!clearValue && cell._richText && !cell.isFormula && typeof cell._editVal !== 'number') {
        this.input.innerHTML = richTextToHTML(cell._richText, cell.font, zoom);
    } else {
        this.input.innerText = clearValue ? '' : getFormulaBarValue(sheet, cell);
    }
    this.currentCell = { ...this.activeSheet.activeCell }; // Prevent async drift
    this.activeSheet.cancelBrush(); // Exit format brush (with toolbar refresh)

    if (!focusEditor || !this.inputEditing) return;

    this.input.focus();
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.selectNodeContents(this.input);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
}

// Update editor position (after zoom or view changes)
export function updateInputPosition() {
    if (!this.inputEditing || !this.currentCell) return;
    const { c, r } = this.currentCell;
    const cell = this.activeSheet.getCell(r, c);
    const cInfo = this.activeSheet.getCellInViewInfo(r, c);
    if (!cInfo) return;
    const zoom = this.activeSheet.zoom ?? 1;
    const baseFontSize = (cell.font.size ?? 11) * 1.333;
    const cellW = cInfo.w * zoom;
    const cellH = cInfo.h * zoom;
    this.input.style.left = `${cInfo.x * zoom}px`;
    this.input.style.top = `${cInfo.y * zoom}px`;
    this.input.style.minWidth = `${Math.max(1, cellW - 1)}px`;
    this.input.style.minHeight = `${Math.max(1, cellH - 1)}px`;
    this.input.style.fontSize = `${baseFontSize * zoom}px`;
    this.input.style.lineHeight = `${baseFontSize * 1.35 * zoom}px`;
}

// Input value update
export function updInputValue() {
    const { c, r } = this.currentCell;
    const sheet = this.activeSheet;
    const cell = sheet.getCell(r, c);

    // 富文本提交路径（数字内容不支持富文本，与 Excel 一致）
    const plainText = this.input.innerText.trim();
    const isNumeric = !isNaN(plainText) && plainText.length <= 15 && plainText !== '';
    if (!isNumeric && hasRichContent(this.input) && !cell.isFormula) {
        const runs = htmlToRichText(this.input, cell.font);
        if (runs && runs.length > 0) {
            cell.richText = runs;
            if (!cell.validData && cell._dataValidation?.showErrorMessage) {
                alert(_trValidationError(this.SN, cell._dataValidation.error));
                this.inputEditing = false;
                this.input.innerHTML = "";
                cell.richText = null;
                return;
            }
            this.inputEditing = false;
            this.input.innerHTML = "";
            if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
                this.SN.DependencyGraph.recalculateVolatile();
            }
            this.r();
            return;
        }
    }

    // 纯文本提交路径
    const inputValue = (isNaN(plainText) || plainText.length > 15 || plainText == '') ? plainText : Number(plainText);
    const old = cell.editVal;
    cell.editVal = inputValue;
    if (!cell.validData && cell._dataValidation.showErrorMessage) {
        alert(_trValidationError(this.SN, cell._dataValidation.error));
        this.inputEditing = false;
        this.input.innerHTML = "";
        return cell.editVal = old;
    }
    this.inputEditing = false;
    this.input.innerHTML = "";

    // 触发 Volatile 函数重算（RAND, TODAY, NOW 等）
    if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
        this.SN.DependencyGraph.recalculateVolatile();
    }

    this.r();
}
