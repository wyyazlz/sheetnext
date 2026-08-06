// 辅助函数：颜色插值
export function interpolateColor(color1, color2, ratio) {
    const r1 = parseInt(color1.substr(1, 2), 16);
    const g1 = parseInt(color1.substr(3, 2), 16);
    const b1 = parseInt(color1.substr(5, 2), 16);

    const r2 = parseInt(color2.substr(1, 2), 16);
    const g2 = parseInt(color2.substr(3, 2), 16);
    const b2 = parseInt(color2.substr(5, 2), 16);

    const r = Math.round(r1 + (r2 - r1) * ratio);
    const g = Math.round(g1 + (g2 - g1) * ratio);
    const b = Math.round(b1 + (b2 - b1) * ratio);

    return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

export function shouldHideFormula(sheet, cell) {
    return !!(sheet?.protection?.enabled && cell?.protection?.hidden && cell?.isFormula);
}

function _hasDateTimeValue(date) {
    return !!(date.getHours() || date.getMinutes() || date.getSeconds() || date.getMilliseconds());
}

export function getFormulaBarValue(sheet, cell) {
    if (!cell) return '';
    if (shouldHideFormula(sheet, cell)) return '';
    if (cell._editVal instanceof Date) {
        if (_hasDateTimeValue(cell._editVal)) return cell._editVal.toLocaleString();
        if (cell.type === 'time') return cell._editVal.toLocaleTimeString();
        return cell._editVal.toLocaleDateString();
    }
    return cell.editVal?.toString() ?? '';
}

function getRangeAxisPanes(layouts, start, end) {
    let frozen = false;
    let scrolling = false;

    for (const layout of layouts) {
        if (layout.index < start || layout.index > end) continue;
        if (layout.frozen) frozen = true;
        else scrolling = true;
        if (frozen && scrolling) break;
    }

    return { frozen, scrolling };
}

/**
 * Resolve the canvas clip occupied by the visible panes of a cell range.
 * @param {Object} canvas - Canvas renderer
 * @param {Object} sheet - Sheet being rendered
 * @param {{s:{r:number,c:number},e:{r:number,c:number}}} area - Cell range
 * @returns {{x:number,y:number,w:number,h:number}|{empty:true}}
 */
export function getRangePaneClip(canvas, sheet, area) {
    const vi = sheet?.vi;
    if (!canvas || !vi || !area?.s || !area?.e) return { empty: true };

    const startRow = Math.min(area.s.r, area.e.r);
    const endRow = Math.max(area.s.r, area.e.r);
    const startCol = Math.min(area.s.c, area.e.c);
    const endCol = Math.max(area.s.c, area.e.c);
    const rowPanes = getRangeAxisPanes(vi.rowLayouts || [], startRow, endRow);
    const colPanes = getRangeAxisPanes(vi.colLayouts || [], startCol, endCol);
    if ((!rowPanes.frozen && !rowPanes.scrolling) || (!colPanes.frozen && !colPanes.scrolling)) {
        return { empty: true };
    }

    const splitX = vi.fw ?? sheet.indexWidth;
    const splitY = vi.fh ?? sheet.headHeight;
    const left = colPanes.frozen ? sheet.indexWidth : splitX;
    const right = colPanes.scrolling ? canvas.viewWidth : splitX;
    const top = rowPanes.frozen ? sheet.headHeight : splitY;
    const bottom = rowPanes.scrolling ? canvas.viewHeight : splitY;
    const x = Math.max(sheet.indexWidth, left);
    const y = Math.max(sheet.headHeight, top);
    const w = Math.min(canvas.viewWidth, right) - x;
    const h = Math.min(canvas.viewHeight, bottom) - y;

    return w > 0 && h > 0 ? { x, y, w, h } : { empty: true };
}

/**
 * Draw within a rectangular pane clip. An empty clip intentionally skips drawing.
 * @param {CanvasRenderingContext2D} ctx - Rendering context
 * @param {{x:number,y:number,w:number,h:number}|{empty:true}|null} clip - Pane clip
 * @param {Function} draw - Drawing callback
 * @returns {void}
 */
export function withPaneClip(ctx, clip, draw) {
    if (clip?.empty) return;
    if (!clip) {
        draw();
        return;
    }

    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.x, clip.y, clip.w, clip.h);
    ctx.clip();
    draw();
    ctx.restore();
}
