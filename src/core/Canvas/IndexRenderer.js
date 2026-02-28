/**
 * 行列标渲染模块
 * 负责渲染行号、列标、全选按钮、高亮区域等
 */

import { syncToolbarState } from '../Layout/ToolbarBuilder.js';
import { getFormulaBarValue } from './helpers.js';

// 获取高亮信息（选中区域的行列）
export function getLight() {
    const sheet = this.activeSheet;
    const areas = sheet.activeAreas;
    const rowsSet = new Set();
    const colsSet = new Set();
    const frames = [];
    areas.forEach(area => {
        const { x, y, w, h, inView } = sheet.getAreaInviewInfo(area);
        frames.push({ x, y, w, h, inView, area });
        // 收集高亮行
        for (let r = area.s.r; r <= area.e.r; r++) {
            if (sheet.vi.rowsIndex.includes(r)) rowsSet.add(r);
        }
        // 收集高亮列
        for (let c = area.s.c; c <= area.e.c; c++) {
            if (sheet.vi.colsIndex.includes(c)) colsSet.add(c);
        }
    });
    this.activeBorderInfo = frames[frames.length - 1]
    return {
        frames,                             // 所有选中区域的位置数据
        lastFrame: this.activeBorderInfo,    // 最后一个框的数据，外部可按原逻辑绘制边框
        rows: Array.from(rowsSet),          // 需要高亮的行索引
        columns: Array.from(colsSet)        // 需要高亮的列索引
    };
}

// 渲染全选按钮
export function rAllBtn() {
    this.bc.beginPath();
    this.bc.strokeStyle = '#ddd';
    this.bc.lineWidth = this.px(1);
    // 在绘制线条和三角形前，先填充背景
    this.bc.fillStyle = '#f0f0f0';
    this.bc.fillRect(0, 0, this.activeSheet.indexWidth, this.activeSheet.headHeight);
    // 绘制直角三角形
    this.bc.beginPath();
    this.bc.moveTo(this.activeSheet.indexWidth / 2, this.activeSheet.headHeight - 5); // 移动到第一个点
    this.bc.lineTo(this.activeSheet.indexWidth - 5, 3); // 画线到第二个点
    this.bc.lineTo(this.activeSheet.indexWidth - 5, this.activeSheet.headHeight - 5); // 画线到第三个点
    this.bc.fillStyle = '#ccc';
    this.bc.fill(); // 填充三角形
}

// 只渲染行标列标（提取出来供截图模式复用）
export function rRowColHeaders(sheet, lightInfo) {
    this.rAllBtn(); // 全选按钮

    let y = sheet.headHeight;
    let x = sheet.indexWidth;
    this.bc.textBaseline = 'middle'
    this.bc.textAlign = 'center'
    this.bc.strokeStyle = '#ccc';
    const lineWidth = this.px(1);
    this.bc.lineWidth = lineWidth;

    // 渲染列头
    this.bc.font = '13px Arial';
    sheet.vi.colsIndex.forEach(c => {
        if (x == 0) return
        const isLight = lightInfo.columns.includes(c)
        let nowCellWidth = sheet.getCol(c).width;

        // 背景
        this.bc.fillStyle = isLight ? '#d9d9d9' : '#f0f0f0';
        this.bc.fillRect(x, 0, nowCellWidth, sheet.headHeight);

        // 边框
        this.bc.beginPath();
        this.bc.strokeStyle = '#ccc';
        const colRect = this.snapRect(x, 0, nowCellWidth, sheet.headHeight);
        this.bc.lineWidth = lineWidth;
        this.bc.strokeRect(colRect.x, colRect.y, colRect.w, colRect.h);

        // 高亮边
        if (isLight) {
            this.bc.beginPath();
            this.bc.lineWidth = 2;
            this.bc.strokeStyle = '#36c';
            const hp = this.halfPixel;
            const yLine = this.snap(sheet.headHeight) + hp;
            this.bc.moveTo(this.snap(x) + hp, yLine);
            this.bc.lineTo(this.snap(x + nowCellWidth) + hp, yLine);
            this.bc.stroke();
        }

        // 文本
        this.bc.beginPath();
        this.bc.lineWidth = lineWidth;
        this.bc.fillStyle = '#333';
        this.bc.fillText(this.Utils.numToChar(c), x + nowCellWidth / 2, sheet.headHeight / 2 + 2);
        x += nowCellWidth;
    })

    // 渲染序号
    this.bc.font = '14.6px Arial';
    sheet.vi.rowsIndex.forEach(r => {
        if (y == 0) return
        const isLight = lightInfo.rows.includes(r)
        x = sheet.indexWidth;
        let rowHeight = sheet.getRow(r).height;

        // 背景
        this.bc.fillStyle = isLight ? '#d9d9d9' : '#f0f0f0';
        this.bc.fillRect(0, y, sheet.indexWidth, rowHeight);

        // 边框
        this.bc.beginPath();
        this.bc.strokeStyle = '#ccc';
        const rowRect = this.snapRect(0, y, sheet.indexWidth, rowHeight);
        this.bc.lineWidth = lineWidth;
        this.bc.strokeRect(rowRect.x, rowRect.y, rowRect.w, rowRect.h);

        // 高亮边
        if (isLight) {
            this.bc.beginPath();
            this.bc.lineWidth = 2;
            this.bc.strokeStyle = '#36c';
            const hp = this.halfPixel;
            const xLine = this.snap(sheet.indexWidth) + hp;
            this.bc.moveTo(xLine, this.snap(y) + hp);
            this.bc.lineTo(xLine, this.snap(y + rowHeight) + hp);
            this.bc.stroke();
        }

        // 文本
        this.bc.beginPath();
        this.bc.lineWidth = lineWidth;
        this.bc.fillStyle = '#333';
        this.bc.fillText(r + 1, sheet.indexWidth / 2, y + rowHeight / 2 + 2);
        y += sheet.getRow(r).height
    })
}

// 合并重叠区间
function mergeRanges(ranges, key1, key2) {
    if (!ranges.length) return [];
    ranges.sort((a, b) => a[key1] - b[key1]);
    const merged = [{ ...ranges[0] }];
    for (let i = 1; i < ranges.length; i++) {
        const last = merged[merged.length - 1];
        const curr = ranges[i];
        if (curr[key1] <= last[key2]) {
            last[key2] = Math.max(last[key2], curr[key2]);
        } else {
            merged.push({ ...curr });
        }
    }
    return merged;
}

function buildStartMap(indices, startCoord, sizeGetter) {
    const starts = new Map();
    let pos = startCoord;
    indices.forEach(index => {
        starts.set(index, pos);
        pos += sizeGetter(index);
    });
    return { starts, end: pos };
}

function isRowsHiddenBetween(sheet, start, end) {
    if (start > end) return true;
    for (let r = start; r <= end; r++) {
        if (!sheet.getRow(r).hidden) return false;
    }
    return true;
}

function isColsHiddenBetween(sheet, start, end) {
    if (start > end) return true;
    for (let c = start; c <= end; c++) {
        if (!sheet.getCol(c).hidden) return false;
    }
    return true;
}

function resolveBoundaryPosition(sheet, index, indices, starts, startCoord, endCoord, axis) {
    if (!Number.isFinite(index) || !indices.length) return null;
    const first = indices[0];
    const last = indices[indices.length - 1];

    if (index <= first) {
        if (index === first) return startCoord;
        return axis === 'row'
            ? (isRowsHiddenBetween(sheet, index, first - 1) ? startCoord : null)
            : (isColsHiddenBetween(sheet, index, first - 1) ? startCoord : null);
    }

    for (const current of indices) {
        if (current >= index) return starts.get(current) ?? null;
    }

    if (index === last + 1) return endCoord;
    if (index > last + 1) {
        return axis === 'row'
            ? (isRowsHiddenBetween(sheet, last + 1, index - 1) ? endCoord : null)
            : (isColsHiddenBetween(sheet, last + 1, index - 1) ? endCoord : null);
    }

    return null;
}

// 绘制高亮条（行列背景，跳过选中区域）
export function rHighlightBars(sheet, lightInfo) {
    const color = this.highlightColor;
    if (!color) return;

    const { indexWidth, headHeight } = sheet;
    const frames = lightInfo.frames;
    this.bc.fillStyle = color;

    // 绘制行高亮条
    lightInfo.rows.forEach(r => {
        let y = headHeight;
        for (const ri of sheet.vi.rowsIndex) {
            if (ri === r) break;
            y += sheet.getRow(ri).height;
        }
        const h = sheet.getRow(r).height;

        // 收集并合并跳过区间
        const skipRanges = frames
            .filter(f => f.inView && r >= f.area.s.r && r <= f.area.e.r)
            .map(f => ({ x1: f.x, x2: f.x + f.w }));
        const merged = mergeRanges(skipRanges, 'x1', 'x2');

        // 分段绘制
        let x = indexWidth;
        for (const skip of merged) {
            if (skip.x1 > x) this.bc.fillRect(x, y, skip.x1 - x, h);
            x = skip.x2;
        }
        if (x < this.viewWidth) this.bc.fillRect(x, y, this.viewWidth - x, h);
    });

    // 绘制列高亮条
    lightInfo.columns.forEach(c => {
        let x = indexWidth;
        for (const ci of sheet.vi.colsIndex) {
            if (ci === c) break;
            x += sheet.getCol(ci).width;
        }
        const w = sheet.getCol(c).width;

        // 收集并合并跳过区间
        const skipRanges = frames
            .filter(f => f.inView && c >= f.area.s.c && c <= f.area.e.c)
            .map(f => ({ y1: f.y, y2: f.y + f.h }));
        const merged = mergeRanges(skipRanges, 'y1', 'y2');

        // 分段绘制
        let y = headHeight;
        for (const skip of merged) {
            if (skip.y1 > y) this.bc.fillRect(x, y, w, skip.y1 - y);
            y = skip.y2;
        }
        if (y < this.viewHeight) this.bc.fillRect(x, y, w, this.viewHeight - y);
    });
}

// 绘制分页符辅助线（类似 Excel 分页线）
export function rPageBreaks(sheet) {
    if (!sheet.showPageBreaks) return;
    const guides = this.SN.Print?.getPageBreakGuides?.(sheet);
    if (!guides?.range) return;

    const rowIndices = sheet.vi?.rowsIndex || [];
    const colIndices = sheet.vi?.colsIndex || [];
    if (!rowIndices.length || !colIndices.length) return;

    const rowMap = buildStartMap(rowIndices, sheet.headHeight, r => sheet.getRow(r).height);
    const colMap = buildStartMap(colIndices, sheet.indexWidth, c => sheet.getCol(c).width);

    const xLeft = resolveBoundaryPosition(sheet, guides.range.s.c, colIndices, colMap.starts, sheet.indexWidth, colMap.end, 'col');
    const xRight = resolveBoundaryPosition(sheet, guides.range.e.c + 1, colIndices, colMap.starts, sheet.indexWidth, colMap.end, 'col');
    const yTop = resolveBoundaryPosition(sheet, guides.range.s.r, rowIndices, rowMap.starts, sheet.headHeight, rowMap.end, 'row');
    const yBottom = resolveBoundaryPosition(sheet, guides.range.e.r + 1, rowIndices, rowMap.starts, sheet.headHeight, rowMap.end, 'row');

    const spanX1 = Math.max(sheet.indexWidth, xLeft ?? sheet.indexWidth);
    const spanX2 = Math.min(this.viewWidth, xRight ?? this.viewWidth);
    const spanY1 = Math.max(sheet.headHeight, yTop ?? sheet.headHeight);
    const spanY2 = Math.min(this.viewHeight, yBottom ?? this.viewHeight);
    if (spanX2 <= spanX1 || spanY2 <= spanY1) return;

    const autoStyle = { color: '#7f9bff', width: 1, dash: [this.px(6), this.px(4)] };
    const manualStyle = { color: '#2f62ff', width: 1.2, dash: [] };
    const borderStyle = { color: '#2f62ff', width: 1.2, dash: [] };
    const hp = this.halfPixel;

    const drawLine = (x1, y1, x2, y2, style) => {
        this.bc.beginPath();
        this.bc.strokeStyle = style.color;
        this.bc.lineWidth = this.px(style.width);
        this.bc.setLineDash(style.dash || []);
        if (Math.abs(x1 - x2) < 0.0001) {
            const x = this.snap(x1) + hp;
            this.bc.moveTo(x, this.snap(y1));
            this.bc.lineTo(x, this.snap(y2));
        } else {
            const y = this.snap(y1) + hp;
            this.bc.moveTo(this.snap(x1), y);
            this.bc.lineTo(this.snap(x2), y);
        }
        this.bc.stroke();
    };

    guides.rowBreaks.forEach(item => {
        const y = resolveBoundaryPosition(sheet, item.index, rowIndices, rowMap.starts, sheet.headHeight, rowMap.end, 'row');
        if (y == null) return;
        drawLine(spanX1, y, spanX2, y, item.manual ? manualStyle : autoStyle);
    });
    guides.colBreaks.forEach(item => {
        const x = resolveBoundaryPosition(sheet, item.index, colIndices, colMap.starts, sheet.indexWidth, colMap.end, 'col');
        if (x == null) return;
        drawLine(x, spanY1, x, spanY2, item.manual ? manualStyle : autoStyle);
    });

    if (xLeft != null) drawLine(xLeft, spanY1, xLeft, spanY2, borderStyle);
    if (xRight != null) drawLine(xRight, spanY1, xRight, spanY2, borderStyle);
    if (yTop != null) drawLine(spanX1, yTop, spanX2, yTop, borderStyle);
    if (yBottom != null) drawLine(spanX1, yBottom, spanX2, yBottom, borderStyle);

    this.bc.setLineDash([]);
}

// 渲染活动框和选区
export function rIndex(sheet, isScreenshot = false, renderOptions = {}) {

    // 截图模式不清除buffer
    if (!isScreenshot) {
        this.clearBuffer(); // 清除缓冲画布
    }

    const lightInfo = this.getLight()
    const selectionVisible = sheet.protection.isSelectionVisible();
    const hideSelectionFrame = !!renderOptions.hideSelectionFrame;

    // 绘制高亮条（在选区之前，作为底层背景）
    if (!isScreenshot && selectionVisible) {
        this.rHighlightBars(sheet, lightInfo);
    }

    if (!isScreenshot) {
        this.rPageBreaks(sheet);
    }

    // 绘制在区域中的选中的区域，将其背景设为#ddd
    if (!isScreenshot && selectionVisible) {
        lightInfo.frames.forEach((frame) => {
            if (!frame.inView) return;
            this.bc.fillStyle = 'rgba(99, 99, 99, 0.2)';
            this.bc.fillRect(frame.x + 1.5, frame.y + 1.5, frame.w - 3, frame.h - 3); // 边之间留了一些小空隙
        });
    }

    // 截图模式跳过行标列标渲染（图表绘制后会统一绘制）
    if (!isScreenshot) {
        this.rRowColHeaders(sheet, selectionVisible ? lightInfo : { rows: [], columns: [] }); // 渲染行标列标
    }

    const lf = lightInfo.lastFrame
    if (selectionVisible && lf && lf.inView && !(isScreenshot && hideSelectionFrame)) {
        let { x, y, w, h } = lf
        // 渲染当前活动的单元格
        const activeCellInfo = sheet.getCellInViewInfo(sheet.activeCell.r, sheet.activeCell.c);
        if (activeCellInfo.inView && !isScreenshot) this.bc.clearRect(activeCellInfo.x + 1, activeCellInfo.y + 1, activeCellInfo.w - 2, activeCellInfo.h - 2);
        const hp = this.halfPixel;
        const rect = this.snapRect(x, y, w, h);
        this.bc.strokeStyle = '#36c';
        this.bc.lineWidth = 1.5;  // 使用固定 CSS 像素，保持视觉一致性
        this.bc.strokeRect(rect.x - hp, rect.y - hp, rect.w + hp * 2, rect.h + hp * 2);
        // 右下角小方块
        const size = 7;
        // 将小方块中心点设在 (x+w, y+h) 处
        const cx = x + w;
        const cy = y + h;
        const px = cx - size / 2;
        const py = cy - size / 2;
        this.bc.fillStyle = '#36c';
        this.bc.fillRect(px, py, size, size);
        this.bc.strokeStyle = '#ffffff';
        this.bc.lineWidth = 2;  // 使用固定 CSS 像素
        this.bc.strokeRect(px, py, size, size);
        this.bc.beginPath();
        this.bc.lineWidth = this.px(1);
    }

    // 截图模式不操作显示画布
    if (!isScreenshot) {
        this.ctx.drawImage(this.showLayerKZ, 0, 0); // 将缓冲区内容复制到显示画布上
        this.ctx.drawImage(this.buffer, 0, 0); // 将缓冲区内容复制到显示画布上

        // 绘制水印
        this.SN.License?._drawWatermark(this.ctx, this.showLayer.width, this.showLayer.height, this.dpr);

        // 绘制追踪箭头
        this.renderTraceArrows(sheet);

        const nowRCount = ++this.rCount
        this.rDrawing(nowRCount); // 图纸渲染
        this.rDom()
    }
}

// dom渲染
export function rDom() {
    const sheet = this.activeSheet;
    if (!sheet.protection.isSelectionVisible()) {
        this.areaInput.value = '';
        this.formulaBar.value = '';
        this.SN.Layout?.closePivotPanel?.();
        return;
    }
    const { r, c } = sheet.activeCell;

    // 区域显示
    const l = sheet.activeAreas;
    if (l.length > 0 && (l[l.length - 1].s.r != l[l.length - 1].e.r || l[l.length - 1].s.c != l[l.length - 1].e.c)) {
        this.areaInput.value = this.Utils.rangeNumToStr(l[l.length - 1])
    } else if (Object.keys(sheet.activeCell).length > 0) {
        this.areaInput.value = this.Utils.numToChar(c) + (r + 1)
    }

    // formulaBar内容
    const cell = sheet.getCell(r, c);
    this.formulaBar.value = getFormulaBarValue(sheet, cell);

    // 同步工具栏状态（纯状态差量对比）
    syncToolbarState(this.SN.containerDom, cell, this.SN.Layout?._toolbarBuilder?.menuConfig);

    // 更新上下文工具栏（表格、图表等）
    this._updateContextualToolbar(sheet, r, c);
}

// 更新上下文工具栏
export function _updateContextualToolbar(sheet, r, c) {
    if (!this.SN.Layout?.updateContextualToolbar) return;

    // 检查当前单元格是否在表格中
    const table = sheet.Table?.getTableAt(r, c);
    if (table) {
        this.SN.Layout.updateContextualToolbar({ type: 'table', data: table });
        return;
    }

    // 离开特殊区域，隐藏上下文菜单
    // 注意：透视表面板的关闭在 InteractionHandler.baseMousedown 中处理
    this.SN.Layout.updateContextualToolbar(null);
}

/**
 * 获取包含指定单元格的透视表
 * @param {Sheet} sheet - 工作表
 * @param {number} row - 行索引
 * @param {number} col - 列索引
 * @returns {PivotTable|null}
 */
function getPivotTableAt(sheet, row, col) {
    if (!sheet.PivotTable || sheet.PivotTable.size === 0) return null;

    for (const pt of sheet.PivotTable.getAll()) {
        if (isPivotTableContainsCell(pt, row, col)) {
            return pt;
        }
    }
    return null;
}

/**
 * 检查透视表是否包含指定单元格
 * @param {PivotTable} pt - 透视表
 * @param {number} row - 行索引
 * @param {number} col - 列索引
 * @returns {boolean}
 */
function isPivotTableContainsCell(pt, row, col) {
    const ref = pt.location?.ref;
    if (!ref) return false;

    // 解析区域引用 (如 "A1:E10")
    const range = parseSimpleRange(ref);
    if (!range) return false;

    return row >= range.s.r && row <= range.e.r && col >= range.s.c && col <= range.e.c;
}

/**
 * 简单解析区域引用
 * @param {string} ref - 区域引用 (如 "A1:E10")
 * @returns {Object|null} {s:{r,c}, e:{r,c}}
 */
function parseSimpleRange(ref) {
    const parts = ref.split(':');
    const start = parseCellRef(parts[0]);
    const end = parts.length > 1 ? parseCellRef(parts[1]) : start;
    if (!start || !end) return null;
    return { s: start, e: end };
}

/**
 * 解析单元格引用
 * @param {string} cellRef - 单元格引用 (如 "A1")
 * @returns {Object|null} {r, c}
 */
function parseCellRef(cellRef) {
    const match = cellRef.replace(/\$/g, '').match(/^([A-Z]+)(\d+)$/i);
    if (!match) return null;

    const colStr = match[1].toUpperCase();
    const row = parseInt(match[2], 10) - 1;

    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    col -= 1;

    return { r: row, c: col };
}
