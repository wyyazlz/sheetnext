/**
 * Filter icon rendering module.
 * Draws filter dropdown icons in header cells.
 */

/** @param {Sheet} sheet */
export function rFilterIcons(sheet) {
    const autoFilter = sheet.AutoFilter;
    this._filterIconPositions = [];

    const scopes = autoFilter?.getEnabledScopes?.();
    if (Array.isArray(scopes) && scopes.length > 0) {
        const colPositions = new Map();
        let colX = sheet.indexWidth;
        for (const c of sheet.vi.colsIndex) {
            const colWidth = sheet.getCol(c).width;
            colPositions.set(c, { x: colX, w: colWidth });
            colX += colWidth;
        }

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

export function clearFilterIconPositions() {
    this._filterIconPositions = [];
}

/** @param {number} x @param {number} y @returns {{colIndex:number, scopeId:string}|null} */
export function getFilterIconAtPosition(x, y) {
    if (!this._filterIconPositions || this._filterIconPositions.length === 0) {
        return null;
    }

    for (const icon of this._filterIconPositions) {
        const padding = 5;
        if (x >= icon.x - padding && x <= icon.x + icon.w + padding &&
            y >= icon.y - padding && y <= icon.y + icon.h + padding) {
            return { colIndex: icon.colIndex, scopeId: icon.scopeId };
        }
    }
    return null;
}

/** @param {number} boxX @param {number} boxY @param {number} boxW @param {number} boxH */
export function getFilterIconLayout(boxX, boxY, boxW, boxH) {
    const padding = 2;
    const iconSize = Math.max(8, Math.min(15, Math.floor(boxH - padding * 2)));
    return {
        iconSize,
        iconX: boxX + Math.max(0, boxW - iconSize - padding),
        iconY: boxY + Math.max(0, boxH - iconSize - padding)
    };
}

function _drawButtonFrame(ctx, iconX, iconY, iconSize, isActive) {
    const gradient = ctx.createLinearGradient(iconX, iconY, iconX, iconY + iconSize);
    gradient.addColorStop(0, '#ffffff');
    gradient.addColorStop(1, isActive ? '#edf3ff' : '#f3f3f3');
    ctx.fillStyle = gradient;
    ctx.fillRect(iconX, iconY, iconSize, iconSize);

    ctx.strokeStyle = isActive ? '#8eaadb' : '#c8c8c8';
    ctx.lineWidth = 1;
    ctx.strokeRect(iconX + 0.5, iconY + 0.5, Math.max(0, iconSize - 1), Math.max(0, iconSize - 1));
}

function _drawDropdownGlyph(ctx, centerX, centerY, color) {
    const size = 4.1;
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(centerX - size, centerY - 1);
    ctx.lineTo(centerX + size, centerY - 1);
    ctx.lineTo(centerX, centerY + size - 1);
    ctx.closePath();
    ctx.fill();
}

function _drawFilterGlyph(ctx, centerX, centerY, color, scale = 1) {
    const topWidth = 7.8 * scale;
    const neckWidth = 2.9 * scale;
    const topY = centerY - 4.2 * scale;
    const midY = centerY - 0.45 * scale;
    const stemBottomY = centerY + 4.2 * scale;

    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(centerX - topWidth / 2, topY);
    ctx.lineTo(centerX + topWidth / 2, topY);
    ctx.lineTo(centerX + neckWidth / 2, midY);
    ctx.lineTo(centerX + neckWidth / 2, stemBottomY);
    ctx.lineTo(centerX - neckWidth / 2, stemBottomY);
    ctx.lineTo(centerX - neckWidth / 2, midY);
    ctx.closePath();
    ctx.fill();
}

function _drawSortGlyph(ctx, centerX, centerY, order, color, scale = 1) {
    const shaftHalf = 1 * scale;
    const top = centerY - 4.6 * scale;
    const bottom = centerY + 4.6 * scale;
    const headSize = 3.6 * scale;
    const shaftTop = order === 'asc' ? top + headSize - 0.25 * scale : top;
    const shaftBottom = order === 'asc' ? bottom : bottom - headSize + 0.25 * scale;

    ctx.fillStyle = color;
    ctx.fillRect(centerX - shaftHalf, shaftTop, shaftHalf * 2, Math.max(0, shaftBottom - shaftTop));

    ctx.beginPath();
    if (order === 'asc') {
        ctx.moveTo(centerX, top);
        ctx.lineTo(centerX - headSize, top + headSize);
        ctx.lineTo(centerX + headSize, top + headSize);
    } else {
        ctx.moveTo(centerX, bottom);
        ctx.lineTo(centerX - headSize, bottom - headSize);
        ctx.lineTo(centerX + headSize, bottom - headSize);
    }
    ctx.closePath();
    ctx.fill();
}

/** @param {CanvasRenderingContext2D} ctx @param {number} iconX @param {number} iconY @param {number} iconSize @param {boolean} hasFilter @param {string|null} sortOrder */
export function drawFilterIcon(ctx, iconX, iconY, iconSize, hasFilter, sortOrder) {
    const isActive = !!(hasFilter || sortOrder);
    const centerX = iconX + iconSize / 2;
    const centerY = iconY + iconSize / 2;
    const primaryColor = isActive ? '#355db3' : '#666666';
    const secondaryColor = '#5b7fcb';

    _drawButtonFrame(ctx, iconX, iconY, iconSize, isActive);

    if (hasFilter && sortOrder) {
        _drawFilterGlyph(ctx, centerX - iconSize * 0.16, centerY, primaryColor, 0.78);
        _drawSortGlyph(ctx, centerX + iconSize * 0.18, centerY, sortOrder, secondaryColor, 0.68);
        return;
    }

    if (hasFilter) {
        _drawFilterGlyph(ctx, centerX, centerY, primaryColor, 0.92);
        return;
    }

    if (sortOrder === 'asc' || sortOrder === 'desc') {
        _drawSortGlyph(ctx, centerX, centerY, sortOrder, primaryColor, 0.92);
        return;
    }

    _drawDropdownGlyph(ctx, centerX, centerY + 0.5, primaryColor);
}
