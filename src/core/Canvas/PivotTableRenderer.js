/**
 * 透视表渲染模块
 * 负责渲染空透视表的占位提示区域
 * 仅渲染视图内的透视表占位
 */

// 占位区域配置
const PLACEHOLDER_COLS = 3;
const PLACEHOLDER_ROWS = 18;

// 图标 SVG 路径（透视表图标）
const PIVOT_ICON_PATH = new Path2D();
// 简化的表格图标
PIVOT_ICON_PATH.moveTo(0, 0);
PIVOT_ICON_PATH.lineTo(24, 0);
PIVOT_ICON_PATH.lineTo(24, 24);
PIVOT_ICON_PATH.lineTo(0, 24);
PIVOT_ICON_PATH.closePath();
PIVOT_ICON_PATH.moveTo(8, 0);
PIVOT_ICON_PATH.lineTo(8, 24);
PIVOT_ICON_PATH.moveTo(16, 0);
PIVOT_ICON_PATH.lineTo(16, 24);
PIVOT_ICON_PATH.moveTo(0, 8);
PIVOT_ICON_PATH.lineTo(24, 8);
PIVOT_ICON_PATH.moveTo(0, 16);
PIVOT_ICON_PATH.lineTo(24, 16);

/**
 * 渲染透视表占位区域
 * @param {Sheet} sheet - 工作表实例
 */
export function rPivotTablePlaceholders(sheet) {
    if (!sheet.PivotTable || sheet.PivotTable.size === 0) return;

    const ctx = this.bc;
    const pixelRatio = this.pixelRatio;

    // 遍历所有透视表
    for (const pt of sheet.PivotTable.getAll()) {
        // 只渲染空透视表（没有字段的）
        if (!isPivotTableEmpty(pt)) continue;

        // 获取透视表起始位置
        const startRef = pt.location.ref.split(':')[0];
        const startCell = this.SN.Utils.cellStrToNum ?
            this.SN.Utils.cellStrToNum(startRef) :
            { r: 0, c: 0 };

        // 计算结束位置
        const endRow = startCell.r + PLACEHOLDER_ROWS - 1;
        const endCol = startCell.c + PLACEHOLDER_COLS - 1;

        // 更新 location.ref 以包含整个占位区域（用于菜单检测）
        const endCellStr = this.SN.Utils.cellNumToStr ?
            this.SN.Utils.cellNumToStr({ r: endRow, c: endCol }) :
            `${String.fromCharCode(65 + endCol)}${endRow + 1}`;
        pt.location.ref = `${startRef}:${endCellStr}`;

        // 计算占位区域在视图中的位置
        const areaInfo = getPlaceholderAreaInfo(sheet, startCell.r, startCell.c);
        if (!areaInfo || !areaInfo.inView) continue;

        // 保存透视表的占位区域信息（用于鼠标检测）
        pt._placeholderArea = {
            x: areaInfo.x,
            y: areaInfo.y,
            w: areaInfo.w,
            h: areaInfo.h,
            startRow: startCell.r,
            startCol: startCell.c
        };

        // 渲染占位区域
        renderPlaceholder(ctx, pt, areaInfo, sheet, pixelRatio, this.SN);
    }
}

/**
 * 判断透视表是否为空
 */
function isPivotTableEmpty(pt) {
    return pt.rowFields.length === 0 &&
        pt.colFields.length === 0 &&
        pt.dataFields.length === 0 &&
        pt.pageFields.length === 0;
}

/**
 * 获取占位区域在视图中的信息（完整区域，不裁剪）
 */
function getPlaceholderAreaInfo(sheet, startRow, startCol) {
    const endRow = startRow + PLACEHOLDER_ROWS - 1;
    const endCol = startCol + PLACEHOLDER_COLS - 1;

    const viewCol = sheet.vi.colsIndex;
    const viewRow = sheet.vi.rowsIndex;

    // 检查区域是否完全在视图外
    if (endRow < viewRow[0] || startRow > viewRow[viewRow.length - 1] ||
        endCol < viewCol[0] || startCol > viewCol[viewCol.length - 1]) {
        return { inView: false };
    }

    // 计算完整区域的位置（可能有负值，表示部分在视图外）
    let x = sheet.indexWidth;
    let y = sheet.headHeight;

    // 计算 x 坐标：从第一个可见列开始，减去滚出部分
    for (let c = viewCol[0]; c < startCol; c++) {
        x += sheet.getCol(c).width;
    }
    // 如果起始列在可见区域之前，x 需要减去
    for (let c = startCol; c < viewCol[0]; c++) {
        x -= sheet.getCol(c).width;
    }

    // 计算 y 坐标：从第一个可见行开始，减去滚出部分
    for (let r = viewRow[0]; r < startRow; r++) {
        y += sheet.getRow(r).height;
    }
    // 如果起始行在可见区域之前，y 需要减去
    for (let r = startRow; r < viewRow[0]; r++) {
        y -= sheet.getRow(r).height;
    }

    // 计算完整区域的宽高
    let w = 0;
    let h = 0;
    for (let c = startCol; c <= endCol; c++) {
        w += sheet.getCol(c).width;
    }
    for (let r = startRow; r <= endRow; r++) {
        h += sheet.getRow(r).height;
    }

    return { x, y, w, h, inView: true };
}

/**
 * 渲染单个透视表占位区域
 */
function renderPlaceholder(ctx, pt, area, sheet, pixelRatio, SN) {
    const { x, y, w, h } = area;
    const snap = (value) => Math.round(value * pixelRatio) / pixelRatio;
    const px = (value) => value / pixelRatio;

    ctx.save();

    // 设置裁剪区域：不绘制到行标和列标区域
    const clipX = Math.max(x, sheet.indexWidth);
    const clipY = Math.max(y, sheet.headHeight);
    const clipW = w - (clipX - x);
    const clipH = h - (clipY - y);

    if (clipW <= 0 || clipH <= 0) {
        ctx.restore();
        return;
    }

    ctx.beginPath();
    ctx.rect(clipX, clipY, clipW, clipH);
    ctx.clip();

    // 1. 绘制背景
    ctx.fillStyle = '#f5f5f5';
    ctx.fillRect(x, y, w, h);

    // 2. 绘制虚线边框
    ctx.strokeStyle = '#b4b4b4';
    ctx.lineWidth = px(1);
    ctx.setLineDash([px(4), px(2)]);
    const rectX1 = snap(x);
    const rectY1 = snap(y);
    const rectX2 = snap(x + w);
    const rectY2 = snap(y + h);
    ctx.strokeRect(rectX1, rectY1, rectX2 - rectX1, rectY2 - rectY1);
    ctx.setLineDash([]);

    // 3. 计算内容区域中心（基于完整区域）
    const centerX = x + w / 2;
    const centerY = y + h / 2;

    // 4. 绘制图标（在上方）
    const iconSize = 32;
    const iconX = centerX - iconSize / 2;
    const iconY = centerY - 60;

    ctx.save();
    ctx.translate(iconX, iconY);
    ctx.scale(iconSize / 24, iconSize / 24);
    ctx.strokeStyle = '#999';
    ctx.lineWidth = 1.5;
    ctx.stroke(PIVOT_ICON_PATH);
    ctx.restore();

    // 5. 绘制透视表名称
    ctx.fillStyle = '#333';
    ctx.font = 'bold 13px "Microsoft YaHei", sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(pt.name, centerX, centerY - 15);

    // 6. 绘制提示语
    ctx.fillStyle = '#666';
    ctx.font = '12px "Microsoft YaHei", sans-serif';
    ctx.fillText(SN.t('core.canvas.pivotTableRenderer.placeholder.generateFrom'), centerX, centerY + 10);
    ctx.fillText(SN.t('core.canvas.pivotTableRenderer.placeholder.selectFields'), centerX, centerY + 28);

    // 7. 绘制分隔线
    ctx.strokeStyle = '#ccc';
    ctx.lineWidth = px(1);
    ctx.beginPath();
    const sepY = snap(centerY + 50);
    ctx.moveTo(snap(centerX - 80), sepY);
    ctx.lineTo(snap(centerX + 80), sepY);
    ctx.stroke();

    // 8. 绘制点击提示
    ctx.fillStyle = '#888';
    ctx.font = '11px "Microsoft YaHei", sans-serif';
    ctx.fillText(SN.t('core.canvas.pivotTableRenderer.placeholder.clickToUse'), centerX, centerY + 70);

    ctx.restore();
}

/**
 * 检测鼠标是否在透视表占位区域内
 * @param {Sheet} sheet - 工作表实例
 * @param {number} x - 鼠标X坐标
 * @param {number} y - 鼠标Y坐标
 * @returns {PivotTable|null} 如果在区域内返回透视表实例，否则返回null
 */
export function getPivotTableAtPosition(sheet, x, y) {
    if (!sheet.PivotTable || sheet.PivotTable.size === 0) return null;

    for (const pt of sheet.PivotTable.getAll()) {
        // 只检测空透视表的占位区域
        if (!isPivotTableEmpty(pt)) continue;

        const area = pt._placeholderArea;
        if (!area) continue;

        // 检测点是否在区域内
        if (x >= area.x && x <= area.x + area.w &&
            y >= area.y && y <= area.y + area.h) {
            return pt;
        }
    }

    return null;
}
