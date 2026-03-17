/**
 * Filter Icon Rendering Module
 * Responsible for drawing the filter drop-down icon in the table header cell
 */

/**
 * Filter Icon Rendering Module
 * Responsible for drawing the filter drop-down icon in the table header cell
 */

/**
 * Render Filter Icon
 * @param {Sheet} sheet - Sheet Objects
 */
export function rFilterIcons(sheet) {
    const autoFilter = sheet.AutoFilter;
    // 清除上一次的图标位置信息
    this._filterIconPositions = [];

    const scopes = autoFilter?.getEnabledScopes?.();
    if (Array.isArray(scopes) && scopes.length > 0) {

    // 预计算可视区列的 x 位置，避免每个 scope 重算
    const colPositions = new Map();
    let colX = sheet.indexWidth;
    for (const c of sheet.vi.colsIndex) {
        const colWidth = sheet.getCol(c).width;
        colPositions.set(c, { x: colX, w: colWidth });
        colX += colWidth;
    }

    // 预计算可视区行的 y 位置
    const rowPositions = new Map();
    let rowY = sheet.headHeight;
    for (const r of sheet.vi.rowsIndex) {
        const rowHeight = sheet.getRow(r).height;
        rowPositions.set(r, { y: rowY, h: rowHeight });
        rowY += rowHeight;
    }

        scopes.forEach(scope => {
            const range = scope.range;
            const headerRow = scope.headerRow;
            const rowPos = rowPositions.get(headerRow);
            if (!range || !rowPos) return;

            for (const c of sheet.vi.colsIndex) {
                if (c < range.s.c || c > range.e.c) continue;
                const colPos = colPositions.get(c);
                if (!colPos) continue;

                const hasFilter = autoFilter.hasFilter(c, scope.scopeId);
                const sortOrder = autoFilter.getSortOrder(c, scope.scopeId);
                const { iconSize, iconX, iconY } = getFilterIconLayout(colPos.x, rowPos.y, colPos.w, rowPos.h);

                this._filterIconPositions.push({
                    x: iconX,
                    y: iconY,
                    w: iconSize,
                    h: iconSize,
                    colIndex: c,
                    scopeId: scope.scopeId
                });

                drawFilterIcon(this.bc, iconX, iconY, iconSize, hasFilter, sortOrder);
            }
        });
    }

    const cellPositions = this._cellRenderPositions || this._cfCellPositions;
    if (!cellPositions?.size) return;

    cellPositions.forEach(({ x, y, w, h, cellInfo }) => {
        const pivotFilter = cellInfo?._pivotFilter;
        if (!pivotFilter) {
            if (cellInfo?._pivotFilterRect) delete cellInfo._pivotFilterRect;
            return;
        }

        const { iconSize, iconX, iconY } = getFilterIconLayout(x, y, w, h);
        if (iconSize <= 0) {
            if (cellInfo._pivotFilterRect) delete cellInfo._pivotFilterRect;
            return;
        }

        drawFilterIcon(this.bc, iconX, iconY, iconSize, !!pivotFilter.active, null);
        cellInfo._pivotFilterRect = { x: iconX, y: iconY, w: iconSize, h: iconSize };
    });
}

/**
 * Clear filter icon position information
 */
export function clearFilterIconPositions() {
    this._filterIconPositions = [];
}

/**
 * Detect if the click location is on the filter icon
 * @param {number} x - Click on the X coordinates
 * @param {number} y - Click on the Y position
 * @returns {{colIndex:number, scopeId:string}|null}
 */
export function getFilterIconAtPosition(x, y) {
    if (!this._filterIconPositions || this._filterIconPositions.length === 0) {
        return null;
    }

    for (const icon of this._filterIconPositions) {
        // 扩大检测区域以便于点击
        const padding = 5;
        if (x >= icon.x - padding && x <= icon.x + icon.w + padding &&
            y >= icon.y - padding && y <= icon.y + icon.h + padding) {
            return { colIndex: icon.colIndex, scopeId: icon.scopeId };
        }
    }
    return null;
}

export function getFilterIconLayout(boxX, boxY, boxW, boxH) {
    const padding = 2;
    const iconSize = Math.max(8, Math.min(15, Math.floor(boxH - padding * 2)));
    return {
        iconSize,
        iconX: boxX + Math.max(0, boxW - iconSize - padding),
        iconY: boxY + Math.max(0, boxH - iconSize - padding)
    };
}

/**
 * Draw a single filter icon (Excel style drop-down button)
 * @private
 */
/** @param {CanvasRenderingContext2D} ctx @param {number} iconX @param {number} iconY @param {number} iconSize @param {boolean} hasFilter @param {string|null} sortOrder */
/**
 * Draw a single filter icon (Excel style drop-down button)
 * @private
 */

export function drawFilterIcon(ctx, iconX, iconY, iconSize, hasFilter, sortOrder) {
    const isActive = hasFilter || sortOrder; // 有筛选或排序时高亮

    // 绘制白色渐变背景
    const gradient = ctx.createLinearGradient(iconX, iconY, iconX, iconY + iconSize);
    gradient.addColorStop(0, '#ffffff');
    gradient.addColorStop(1, '#f0f0f0');
    ctx.fillStyle = gradient;
    ctx.fillRect(iconX, iconY, iconSize, iconSize);
    // Border is aligned to 0.5px so it matches grid-line thickness.
    ctx.strokeStyle = isActive ? '#4f7cf6' : '#ddd';
    ctx.lineWidth = 1;
    ctx.strokeRect(iconX + 0.5, iconY + 0.5, Math.max(0, iconSize - 1), Math.max(0, iconSize - 1));

    // 绘制三角形
    const cx = iconX + iconSize / 2;
    const cy = iconY + iconSize / 2;
    const triangleSize = 4;

    ctx.fillStyle = isActive ? '#4f7cf6' : '#666';
    ctx.beginPath();

    if (sortOrder === 'asc') {
        // 升序：向上的三角形
        ctx.moveTo(cx - triangleSize, cy + 1);
        ctx.lineTo(cx + triangleSize, cy + 1);
        ctx.lineTo(cx, cy - triangleSize + 1);
    } else if (sortOrder === 'desc') {
        // 降序：向下的三角形
        ctx.moveTo(cx - triangleSize, cy - 1);
        ctx.lineTo(cx + triangleSize, cy - 1);
        ctx.lineTo(cx, cy + triangleSize - 1);
    } else {
        // 无排序：默认向下的三角形
        ctx.moveTo(cx - triangleSize, cy - 1);
        ctx.lineTo(cx + triangleSize, cy - 1);
        ctx.lineTo(cx, cy + triangleSize - 1);
    }

    ctx.closePath();
    ctx.fill();
}
