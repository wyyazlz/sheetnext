// 样式操作模块

import { getPosSvg } from '../assets/borderSvgs.js';
import { applyFormatToSelection, getSelectionFont } from '../core/Canvas/RichTextEditor.js';

// 主题单元格样式生成器
function _generateThemedStyles() {
    const accents = [
        { name: 'accent1', colors: ['#DCE6F1', '#B8CCE4', '#4F81BD'] },
        { name: 'accent2', colors: ['#FDE9D9', '#FCD5B4', '#F79646'] },
        { name: 'accent3', colors: ['#D9D9D9', '#BFBFBF', '#808080'] },
        { name: 'accent4', colors: ['#FFF2CC', '#FFE599', '#FFCC00'] },
        { name: 'accent5', colors: ['#DAEEF3', '#B7DEE8', '#4BACC6'] },
        { name: 'accent6', colors: ['#EBF1DE', '#D8E4BC', '#9BBB59'] },
    ];
    const levels = ['20%', '40%', '60%'];
    const levelTokens = ['20', '40', '60'];
    const result = [];
    for (let lvl = 0; lvl < 3; lvl++) {
        for (const a of accents) {
            result.push({
                name: `${a.name}_${levels[lvl]}`,
                labelKey: `action.style.cellStyle.item.${a.name}_${levelTokens[lvl]}`,
                font: lvl === 2 ? { color: '#FFFFFF' } : null,
                fill: { fgColor: a.colors[lvl] },
            });
        }
    }
    return result;
}

export const CELL_STYLE_GROUP_META = {
    goodBadNeutral: {
        labelKey: 'action.style.cellStyle.group.goodBadNeutral'
    },
    dataAndModel: {
        labelKey: 'action.style.cellStyle.group.dataAndModel'
    },
    titlesAndHeadings: {
        labelKey: 'action.style.cellStyle.group.titlesAndHeadings'
    },
    themedCellStyles: {
        labelKey: 'action.style.cellStyle.group.themedCellStyles'
    },
    numberFormat: {
        labelKey: 'action.style.cellStyle.group.numberFormat'
    }
};

export const CELL_STYLE_PRESETS = {
    goodBadNeutral: [
        { name: 'good', labelKey: 'action.style.cellStyle.item.good', font: { color: '#006100' }, fill: { fgColor: '#C6EFCE' } },
        { name: 'bad', labelKey: 'action.style.cellStyle.item.bad', font: { color: '#9C0006' }, fill: { fgColor: '#FFC7CE' } },
        { name: 'neutral', labelKey: 'action.style.cellStyle.item.neutral', font: { color: '#9C5700' }, fill: { fgColor: '#FFEB9C' } },
    ],
    dataAndModel: [
        { name: 'calculation', labelKey: 'action.style.cellStyle.item.calculation', font: { color: '#FA7D00', bold: true }, fill: { fgColor: '#F2F2F2' }, border: { style: 'thin', color: '#7F7F7F' } },
        { name: 'checkCell', labelKey: 'action.style.cellStyle.item.checkCell', font: { color: '#FFFFFF', bold: true }, fill: { fgColor: '#A5A5A5' }, border: { style: 'double', color: '#3F3F76' } },
        { name: 'explanatory', labelKey: 'action.style.cellStyle.item.explanatory', font: { color: '#7F7F7F', italic: true } },
        { name: 'input', labelKey: 'action.style.cellStyle.item.input', font: { color: '#3F3F76' }, fill: { fgColor: '#FFCC99' }, border: { style: 'thin', color: '#7F7F7F' } },
        { name: 'linkedCell', labelKey: 'action.style.cellStyle.item.linkedCell', font: { color: '#FA7D00' }, border: { style: 'thin', color: '#FF8001', area: 'bottom' } },
        { name: 'note', labelKey: 'action.style.cellStyle.item.note', fill: { fgColor: '#FFFFCC' }, border: { style: 'thin', color: '#B2B2B2' } },
        { name: 'output', labelKey: 'action.style.cellStyle.item.output', font: { color: '#3F3F3F', bold: true }, fill: { fgColor: '#F2F2F2' }, border: { style: 'thin', color: '#3F3F3F' } },
        { name: 'warningText', labelKey: 'action.style.cellStyle.item.warningText', font: { color: '#FF0000' } },
    ],
    titlesAndHeadings: [
        { name: 'title', labelKey: 'action.style.cellStyle.item.title', font: { color: '#44546A', size: 18, bold: true } },
        { name: 'heading1', labelKey: 'action.style.cellStyle.item.heading1', font: { color: '#44546A', size: 15, bold: true }, border: { style: 'thick', color: '#5B9BD5', area: 'bottom' } },
        { name: 'heading2', labelKey: 'action.style.cellStyle.item.heading2', font: { color: '#44546A', size: 13, bold: true }, border: { style: 'thick', color: '#5B9BD5', area: 'bottom' } },
        { name: 'heading3', labelKey: 'action.style.cellStyle.item.heading3', font: { color: '#44546A', bold: true }, border: { style: 'medium', color: '#5B9BD5', area: 'bottom' } },
        { name: 'heading4', labelKey: 'action.style.cellStyle.item.heading4', font: { color: '#44546A', bold: true } },
        { name: 'total', labelKey: 'action.style.cellStyle.item.total', font: { bold: true }, border: { style: 'thin', color: '#5B9BD5', area: 'top' }, border2: { style: 'double', color: '#5B9BD5', area: 'bottom' } },
    ],
    themedCellStyles: _generateThemedStyles(),
    numberFormat: [
        { name: 'comma', labelKey: 'action.style.cellStyle.item.comma', numFmt: '#,##0.00' },
        { name: 'commaNoDecimal', labelKey: 'action.style.cellStyle.item.commaNoDecimal', numFmt: '#,##0' },
        { name: 'currency', labelKey: 'action.style.cellStyle.item.currency', numFmt: '¥#,##0.00' },
        { name: 'currencyNoDecimal', labelKey: 'action.style.cellStyle.item.currencyNoDecimal', numFmt: '¥#,##0' },
        { name: 'percent', labelKey: 'action.style.cellStyle.item.percent', numFmt: '0%' },
    ],
};

/**
 * 检测选区是否为整行/整列选择
 * @param {Object} region - 选区 { s: {r, c}, e: {r, c} }
 * @param {Object} sheet - 工作表
 * @returns {{ isFullRow: boolean, isFullCol: boolean, rows: number[], cols: number[] }}
 */
function detectFullRowCol(region, sheet) {
    const { s, e } = region;
    const isFullRow = s.c === 0 && e.c === sheet.colCount - 1;
    const isFullCol = s.r === 0 && e.r === sheet.rowCount - 1;

    const rows = [];
    const cols = [];

    if (isFullRow) {
        for (let r = s.r; r <= e.r; r++) rows.push(r);
    }
    if (isFullCol) {
        for (let c = s.c; c <= e.c; c++) cols.push(c);
    }

    return { isFullRow, isFullCol, rows, cols };
}

const TOOLBAR_DEFAULTS = {
    font: '#FF0000',
    fill: '#FFFF00',
    border: 'all',
    borderStyle: 'thin'
};

function getToolbarRoot(SN) {
    return SN?.containerDom?.querySelector('.sn-tools') || null;
}

function updateColorButtonEmptyState(SN, type, color) {
    const btn = SN?.containerDom?.querySelector(`.sn-split-btn[data-color-type="${type}"]`);
    if (!btn) return;
    if (color) {
        btn.removeAttribute('data-color-empty');
    } else {
        btn.setAttribute('data-color-empty', 'true');
    }
}

function ensureToolbarDefaults(SN) {
    if (!SN) return;
    if (!SN.toolbarDefaults) {
        SN.toolbarDefaults = { ...TOOLBAR_DEFAULTS };
    } else {
        Object.keys(TOOLBAR_DEFAULTS).forEach(key => {
            if (SN.toolbarDefaults[key] === undefined) SN.toolbarDefaults[key] = TOOLBAR_DEFAULTS[key];
        });
    }

    const root = getToolbarRoot(SN);
    if (root) {
        root.style.setProperty('--sn-font-color', SN.toolbarDefaults.font || 'transparent');
        root.style.setProperty('--sn-fill-color', SN.toolbarDefaults.fill || 'transparent');
    }
    updateColorButtonEmptyState(SN, 'font', SN.toolbarDefaults.font);
    updateColorButtonEmptyState(SN, 'fill', SN.toolbarDefaults.fill);
}

function setToolbarDefaultColor(SN, type, color) {
    ensureToolbarDefaults(SN);
    SN.toolbarDefaults[type] = color || '';
    const root = getToolbarRoot(SN);
    if (root) {
        const varName = type === 'font' ? '--sn-font-color' : '--sn-fill-color';
        root.style.setProperty(varName, color || 'transparent');
    }
    updateColorButtonEmptyState(SN, type, color);
}

function getToolbarDefaultColor(SN, type) {
    ensureToolbarDefaults(SN);
    return SN.toolbarDefaults[type];
}

export function initToolbarDefaults() {
    ensureToolbarDefaults(this.SN);
    updateBorderIcon(this.SN, this.SN.toolbarDefaults.border);
}

function getSplitButtonIcon(SN, toolKey) {
    return SN?.containerDom?.querySelector(`.sn-split-btn[data-tool="${toolKey}"] .sn-svg`) || null;
}

function setSplitButtonIconSvg(SN, toolKey, svgHtml) {
    const icon = getSplitButtonIcon(SN, toolKey);
    if (!icon) return;
    icon.outerHTML = svgHtml;
}

function updateBorderIcon(SN, pos) {
    setSplitButtonIconSvg(SN, 'border', getPosSvg(pos || 'all', 16));
}

export function fillColorBtn(e) {
    const target = e.target;
    const classList = target.classList;
    if (classList.contains('sn-color-item')) {
        const color = target.dataset.color;
        setToolbarDefaultColor(this.SN, 'fill', color);
        this.areasStyle('fill', { fgColor: color });
    } else if (classList.contains('sn-color-no-btn')) {
        setToolbarDefaultColor(this.SN, 'fill', '');
        this.areasStyle('fill', {});
    }
}

export function fillColorChange(e) {
    const color = e.target.value;
    setToolbarDefaultColor(this.SN, 'fill', color);
    this.areasStyle('fill', { fgColor: color });
}

export function borderColorBtn(e) {
    const target = e.target;
    const classList = target.classList;
    if (classList.contains('sn-color-item')) {
        this.SN.activeSheet.areasBorder('all', { color: target.dataset.color });
        this.SN._r();
    } else if (classList.contains('sn-color-no-btn')) {
        this.SN.activeSheet.areasBorder('all', { color: null });
        this.SN._r();
    }
}

export function borderColorChange(e) {
    this.SN.activeSheet.areasBorder('all', { color: e.target.value });
    this.SN._r();
}

export function fontColorBtn(e) {
    const target = e.target;
    const classList = target.classList;
    if (classList.contains('sn-color-item')) {
        const color = target.dataset.color;
        setToolbarDefaultColor(this.SN, 'font', color);
        this.areasStyle('font', { color });
    } else if (classList.contains('sn-color-no-btn')) {
        setToolbarDefaultColor(this.SN, 'font', '');
        this.areasStyle('font', { color: null });
    }
}

export function fontColorChange(e) {
    const color = e.target.value;
    setToolbarDefaultColor(this.SN, 'font', color);
    this.areasStyle('font', { color });
}

export function applyFontColor() {
    const color = getToolbarDefaultColor(this.SN, 'font');
    if (color) {
        this.areasStyle('font', { color });
    } else {
        this.areasStyle('font', { color: null });
    }
}

export function applyFillColor() {
    const color = getToolbarDefaultColor(this.SN, 'fill');
    if (color) {
        this.areasStyle('fill', { fgColor: color });
    } else {
        this.areasStyle('fill', {});
    }
}

const BORDER_POS_MAP = {
    none: { area: 'none', style: null },
    all: { area: 'all' },
    outer: { area: 'out' },
    inner: { area: 'inner' },
    left: { area: 'left' },
    top: { area: 'top' },
    right: { area: 'right' },
    bottom: { area: 'bottom' },
    diagonalDown: { area: 'diagonalDown' },
    diagonalUp: { area: 'diagonalUp' }
};

export function applyBorderPreset(pos) {
    ensureToolbarDefaults(this.SN);
    const preset = pos || this.SN.toolbarDefaults.border || 'all';
    const mapping = BORDER_POS_MAP[preset] || BORDER_POS_MAP.all;
    const style = mapping.style !== undefined ? mapping.style : (this.SN.toolbarDefaults.borderStyle || 'thin');
    this.SN.toolbarDefaults.border = preset;
    updateBorderIcon(this.SN, preset);
    this.SN.activeSheet.areasBorder(mapping.area, style ? { style } : null);
    this.SN._r();
}

export function applyBorderStyle(style) {
    ensureToolbarDefaults(this.SN);
    this.SN.toolbarDefaults.borderStyle = style;
    this.SN.activeSheet.areasBorder('all', { style });
    this.SN._r();
}

export function areasStyle(type, options) {
    // 富文本编辑拦截：编辑中对选中文字应用字体格式
    if (type === 'font' && this.SN.Canvas?.inputEditing) {
        const cv = this.SN.Canvas;
        const cell = this.SN.activeSheet.getCell(cv.currentCell.r, cv.currentCell.c);
        if (!cell.isFormula) {
            for (const [prop, value] of Object.entries(options)) {
                applyFormatToSelection(cv.input, prop, value);
            }
            cv.formulaBar.value = cv.input.innerText;
            return;
        }
    }

    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;

    // 收集需要应用样式的行/列/单元格
    const styledRows = new Set();
    const styledCols = new Set();

    areas.forEach(region => {
        if (typeof region === 'string') region = sheet.rangeStrToNum(region);
        const { isFullRow, isFullCol, rows, cols } = detectFullRowCol(region, sheet);

        if (isFullRow && !isFullCol) {
            // 整行选择：应用到行样式
            rows.forEach(r => styledRows.add(r));
        } else if (isFullCol && !isFullRow) {
            // 整列选择：应用到列样式
            cols.forEach(c => styledCols.add(c));
        }
        // 全选或普通单元格选择：继续下面的单元格遍历
    });

    // 应用行样式
    styledRows.forEach(r => {
        sheet.getRow(r)[type] = options;
    });

    // 应用列样式
    styledCols.forEach(c => {
        sheet.getCol(c)[type] = options;
    });

    // 对非整行整列的区域，应用到单元格
    sheet.eachCells(areas, (r, c, range) => {
        const { isFullRow, isFullCol } = detectFullRowCol({ s: range.s, e: range.e }, sheet);
        // 跳过整行或整列选择（已处理）
        if ((isFullRow && !isFullCol) || (isFullCol && !isFullRow)) return;
        sheet.getCell(r, c)[type] = options;
    });

    this.SN._r();
}

export function fontInversion(key, type) {
    // 富文本编辑拦截
    if (this.SN.Canvas?.inputEditing) {
        const cv = this.SN.Canvas;
        const cell = this.SN.activeSheet.getCell(cv.currentCell.r, cv.currentCell.c);
        if (!cell.isFormula) {
            applyFormatToSelection(cv.input, key, true);
            cv.formulaBar.value = cv.input.innerText;
            return;
        }
    }

    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;
    let n = null;

    // 获取参考值（从第一个选区的第一个目标获取）
    let firstArea = areas[0];
    if (typeof firstArea === 'string') firstArea = sheet.rangeStrToNum(firstArea);
    const { isFullRow, isFullCol } = detectFullRowCol(firstArea, sheet);

    if (isFullRow && !isFullCol) {
        const refRow = sheet.getRow(firstArea.s.r);
        if (key === 'underline') {
            n = refRow.font[key] === type ? '' : type;
        } else {
            n = !refRow.font[key];
        }
    } else if (isFullCol && !isFullRow) {
        const refCol = sheet.getCol(firstArea.s.c);
        if (key === 'underline') {
            n = refCol.font[key] === type ? '' : type;
        } else {
            n = !refCol.font[key];
        }
    } else {
        const refCell = sheet.getCell(firstArea.s.r, firstArea.s.c);
        if (key === 'underline') {
            n = refCell.font[key] === type ? '' : type;
        } else {
            n = !refCell.font[key];
        }
    }

    // 收集需要应用样式的行/列
    const styledRows = new Set();
    const styledCols = new Set();

    areas.forEach(region => {
        if (typeof region === 'string') region = sheet.rangeStrToNum(region);
        const info = detectFullRowCol(region, sheet);

        if (info.isFullRow && !info.isFullCol) {
            info.rows.forEach(r => styledRows.add(r));
        } else if (info.isFullCol && !info.isFullRow) {
            info.cols.forEach(c => styledCols.add(c));
        }
    });

    // 应用行样式
    styledRows.forEach(r => {
        sheet.getRow(r).font = { [key]: n };
    });

    // 应用列样式
    styledCols.forEach(c => {
        sheet.getCol(c).font = { [key]: n };
    });

    // 对非整行整列的区域，应用到单元格
    sheet.eachCells(areas, (r, c, range) => {
        const info = detectFullRowCol({ s: range.s, e: range.e }, sheet);
        if ((info.isFullRow && !info.isFullCol) || (info.isFullCol && !info.isFullRow)) return;
        sheet.getCell(r, c).font = { [key]: n };
    });

    this.SN._r();
}

export function toggleWrapText() {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;
    let wrapText = null;

    // 获取参考值
    let firstArea = areas[0];
    if (typeof firstArea === 'string') firstArea = sheet.rangeStrToNum(firstArea);
    const { isFullRow, isFullCol } = detectFullRowCol(firstArea, sheet);

    if (isFullRow && !isFullCol) {
        wrapText = !sheet.getRow(firstArea.s.r).alignment?.wrapText;
    } else if (isFullCol && !isFullRow) {
        wrapText = !sheet.getCol(firstArea.s.c).alignment?.wrapText;
    } else {
        wrapText = !sheet.getCell(firstArea.s.r, firstArea.s.c).alignment?.wrapText;
    }

    // 收集行/列
    const styledRows = new Set();
    const styledCols = new Set();

    areas.forEach(region => {
        if (typeof region === 'string') region = sheet.rangeStrToNum(region);
        const info = detectFullRowCol(region, sheet);

        if (info.isFullRow && !info.isFullCol) {
            info.rows.forEach(r => styledRows.add(r));
        } else if (info.isFullCol && !info.isFullRow) {
            info.cols.forEach(c => styledCols.add(c));
        }
    });

    // 应用行样式
    styledRows.forEach(r => {
        sheet.getRow(r).alignment = { wrapText };
    });

    // 应用列样式
    styledCols.forEach(c => {
        sheet.getCol(c).alignment = { wrapText };
    });

    // 对非整行整列的区域，应用到单元格
    sheet.eachCells(areas, (r, c, range) => {
        const info = detectFullRowCol({ s: range.s, e: range.e }, sheet);
        if ((info.isFullRow && !info.isFullCol) || (info.isFullCol && !info.isFullRow)) return;
        sheet.getCell(r, c).alignment = { wrapText };
    });

    this.SN._r();
}

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

export function changeFontSize(increase = true) {
    const sheet = this.SN.activeSheet;
    const cell = sheet.getCell(sheet.activeCell.r, sheet.activeCell.c);

    // 富文本编辑拦截：从选区获取当前字号
    if (this.SN.Canvas?.inputEditing && !cell.isFormula) {
        const cv = this.SN.Canvas;
        const selFont = getSelectionFont(cv.input, cell.font);
        const current = selFont.size || 11;
        let idx = FONT_SIZES.findIndex(s => s >= current);
        if (idx === -1) idx = increase ? FONT_SIZES.length - 1 : FONT_SIZES.length - 2;
        else idx = increase ? Math.min(idx + 1, FONT_SIZES.length - 1) : Math.max(idx - 1, 0);
        applyFormatToSelection(cv.input, 'size', FONT_SIZES[idx]);
        cv.formulaBar.value = cv.input.innerText;
        return;
    }

    const current = cell.style?.font?.size || 11;
    let idx = FONT_SIZES.findIndex(s => s >= current);
    if (idx === -1) idx = increase ? FONT_SIZES.length - 1 : FONT_SIZES.length - 2;
    else idx = increase ? Math.min(idx + 1, FONT_SIZES.length - 1) : Math.max(idx - 1, 0);
    this.areasStyle('font', { size: FONT_SIZES[idx] });
}

export function changeIndent(delta) {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;

    // 收集行/列
    const styledRows = new Set();
    const styledCols = new Set();

    areas.forEach(region => {
        if (typeof region === 'string') region = sheet.rangeStrToNum(region);
        const info = detectFullRowCol(region, sheet);

        if (info.isFullRow && !info.isFullCol) {
            info.rows.forEach(r => styledRows.add(r));
        } else if (info.isFullCol && !info.isFullRow) {
            info.cols.forEach(c => styledCols.add(c));
        }
    });

    // 应用行样式
    styledRows.forEach(r => {
        const row = sheet.getRow(r);
        const current = parseInt(row.alignment?.indent) || 0;
        const next = Math.max(0, Math.min(250, current + delta));
        row.alignment = { indent: next || undefined };
    });

    // 应用列样式
    styledCols.forEach(c => {
        const col = sheet.getCol(c);
        const current = parseInt(col.alignment?.indent) || 0;
        const next = Math.max(0, Math.min(250, current + delta));
        col.alignment = { indent: next || undefined };
    });

    // 对非整行整列的区域，应用到单元格
    sheet.eachCells(areas, (r, c, range) => {
        const info = detectFullRowCol({ s: range.s, e: range.e }, sheet);
        if ((info.isFullRow && !info.isFullCol) || (info.isFullCol && !info.isFullRow)) return;
        const cell = sheet.getCell(r, c);
        const current = parseInt(cell.alignment?.indent) || 0;
        const next = Math.max(0, Math.min(250, current + delta));
        cell.alignment = { indent: next || undefined };
    });

    this.SN._r();
}

// 文字方向预设值
const TEXT_ROTATION_PRESETS = {
    none: 0,           // 无旋转
    ccw45: 45,         // 逆时针45°
    cw45: 135,         // 顺时针45°
    vertical: 255,     // 竖排
    up90: 90,          // 向上90°
    down90: 180        // 向下90°
};

export function setTextRotation(value) {
    // 支持预设名或数值
    const rotation = TEXT_ROTATION_PRESETS[value] ?? (parseInt(value) || 0);
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;

    // 收集行/列
    const styledRows = new Set();
    const styledCols = new Set();

    areas.forEach(region => {
        if (typeof region === 'string') region = sheet.rangeStrToNum(region);
        const info = detectFullRowCol(region, sheet);

        if (info.isFullRow && !info.isFullCol) {
            info.rows.forEach(r => styledRows.add(r));
        } else if (info.isFullCol && !info.isFullRow) {
            info.cols.forEach(c => styledCols.add(c));
        }
    });

    // 应用行样式
    styledRows.forEach(r => {
        sheet.getRow(r).alignment = { textRotation: rotation || undefined };
    });

    // 应用列样式
    styledCols.forEach(c => {
        sheet.getCol(c).alignment = { textRotation: rotation || undefined };
    });

    // 对非整行整列的区域，应用到单元格
    sheet.eachCells(areas, (r, c, range) => {
        const info = detectFullRowCol({ s: range.s, e: range.e }, sheet);
        if ((info.isFullRow && !info.isFullCol) || (info.isFullCol && !info.isFullRow)) return;
        sheet.getCell(r, c).alignment = { textRotation: rotation || undefined };
    });

    this.SN._r();
}

/**
 * 应用单元格样式预设
 * @param {string} styleName - 样式名称
 */
export function applyCellStyle(styleName) {
    let preset = null;
    for (const styles of Object.values(CELL_STYLE_PRESETS)) {
        preset = styles.find(s => s.name === styleName);
        if (preset) break;
    }
    if (!preset) return;

    const sheet = this.SN.activeSheet;

    if (preset.font) areasStyle.call(this, 'font', preset.font);
    if (preset.fill) areasStyle.call(this, 'fill', preset.fill);

    if (preset.border) {
        sheet.areasBorder(preset.border.area || 'all', {
            style: preset.border.style || 'thin',
            color: preset.border.color
        });
    }
    if (preset.border2) {
        sheet.areasBorder(preset.border2.area || 'all', {
            style: preset.border2.style || 'thin',
            color: preset.border2.color
        });
    }

    if (preset.numFmt) {
        sheet.eachCells(sheet.activeAreas, (r, c) => {
            sheet.getCell(r, c).numFmt = preset.numFmt;
        });
    }

    this.SN._r();
}
