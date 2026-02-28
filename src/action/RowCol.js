// 行列操作模块
// 架构: Core(纯逻辑) -> Adapters(UI适配)

// ============================================================
// Core Operations - 纯业务逻辑，无UI耦合
// 简单操作使用公开属性（自动触发事件和UndoRedo）
// 批量操作（如autoFit）内部处理UndoRedo
// ============================================================

const Core = {
    // 插入行（RowColHandler 自带 UndoRedo）
    insertRows(sheet, startRow, count = 1) {
        sheet.addRows(startRow, count);
    },

    // 插入列（RowColHandler 自带 UndoRedo）
    insertCols(sheet, startCol, count = 1) {
        sheet.addCols(startCol, count);
    },

    // 删除行（RowColHandler 自带 UndoRedo）
    deleteRows(sheet, startRow, count = 1) {
        sheet.delRows(startRow, count);
    },

    // 删除列（RowColHandler 自带 UndoRedo）
    deleteCols(sheet, startCol, count = 1) {
        sheet.delCols(startCol, count);
    },

    // 隐藏行
    hideRows(sheet, startRow, endRow) {
        for (let i = startRow; i <= endRow; i++) {
            sheet.getRow(i).hidden = true;
        }
    },

    // 隐藏列
    hideCols(sheet, startCol, endCol) {
        for (let i = startCol; i <= endCol; i++) {
            sheet.getCol(i).hidden = true;
        }
    },

    // 取消隐藏行
    unhideRows(sheet, startRow, endRow) {
        for (let i = startRow; i <= endRow; i++) {
            sheet.getRow(i).hidden = false;
        }
    },

    // 取消隐藏列
    unhideCols(sheet, startCol, endCol) {
        for (let i = startCol; i <= endCol; i++) {
            sheet.getCol(i).hidden = false;
        }
    },

    // 取消隐藏所有行
    unhideAllRows(sheet) {
        sheet.rows.forEach(row => {
            if (row?.hidden) row.hidden = false;
        });
    },

    // 取消隐藏所有列
    unhideAllCols(sheet) {
        sheet.cols.forEach(col => {
            if (col?.hidden) col.hidden = false;
        });
    },

    // 设置行高
    setRowsHeight(sheet, startRow, endRow, height) {
        for (let i = startRow; i <= endRow; i++) {
            sheet.getRow(i).height = height;
        }
    },

    // 设置列宽
    setColsWidth(sheet, startCol, endCol, width) {
        for (let i = startCol; i <= endCol; i++) {
            sheet.getCol(i).width = width;
        }
    },

    // 自适应行高（稀疏遍历）
    autoFitRowsHeight(sheet, startRow, endRow) {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const results = [];
        const fontCache = new Map();
        const measureCache = new Map();
        const colWidthCache = new Map();

        const getFontInfo = (font) => {
            const size = font.size || 11;
            const name = font.name || '微软雅黑';
            const bold = font.bold ? 'bold' : '';
            const italic = font.italic ? 'italic' : '';
            const key = `${name}|${size}|${bold}|${italic}`;
            let str = fontCache.get(key);
            if (!str) {
                str = `${bold ? 'bold ' : ''}${italic ? 'italic ' : ''}${size}pt ${name}`;
                fontCache.set(key, str);
            }
            return { key, str, size };
        };

        const measureTextWidth = (fontKey, fontStr, text) => {
            const cacheKey = `${fontKey}|${text}`;
            if (measureCache.has(cacheKey)) return measureCache.get(cacheKey);
            ctx.font = fontStr;
            const width = ctx.measureText(text).width;
            measureCache.set(cacheKey, width);
            return width;
        };

        for (let r = startRow; r <= endRow; r++) {
            const row = sheet.rows[r];
            if (!row || !row.cells?.length) continue;
            let maxHeight = sheet.defaultRowHeight;
            let hasContent = false;

            row.cells.forEach((cell, c) => {
                if (!cell) return;
                const value = cell.showVal;
                if (!value) return;
                hasContent = true;

                const font = cell.style?.font || {};
                const { key, str, size } = getFontInfo(font);
                const colWidth = colWidthCache.has(c)
                    ? colWidthCache.get(c)
                    : (sheet.cols[c]?._width ?? sheet.defaultColWidth);
                colWidthCache.set(c, colWidth);

                const lines = String(value).split('\n');
                let totalHeight = 0;
                for (const line of lines) {
                    const width = measureTextWidth(key, str, line);
                    const wrapLines = Math.max(1, Math.ceil(width / Math.max(1, colWidth - 8)));
                    totalHeight += wrapLines * (size * 1.5);
                }
                maxHeight = Math.max(maxHeight, totalHeight + 6);
            });

            if (!hasContent) continue;
            const newHeight = Math.round(maxHeight);
            if (newHeight === row._height) continue;
            results.push({ row, newHeight, oldHeight: row._height });
            row._height = newHeight;
        }

        // 批量 UndoRedo
        if (results.length > 0) {
            sheet.SN.UndoRedo.add({
                undo: () => results.forEach(({ row, oldHeight }) => row._height = oldHeight),
                redo: () => results.forEach(({ row, newHeight }) => row._height = newHeight)
            });
        }
    },

    // 自适应列宽（稀疏遍历）
    autoFitColsWidth(sheet, startCol, endCol) {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        const results = [];
        const fontCache = new Map();
        const measureCache = new Map();
        const maxWidthByCol = new Map();

        const getFontInfo = (font) => {
            const size = font.size || 11;
            const name = font.name || '微软雅黑';
            const bold = font.bold ? 'bold' : '';
            const italic = font.italic ? 'italic' : '';
            const key = `${name}|${size}|${bold}|${italic}`;
            let str = fontCache.get(key);
            if (!str) {
                str = `${bold ? 'bold ' : ''}${italic ? 'italic ' : ''}${size}pt ${name}`;
                fontCache.set(key, str);
            }
            return { key, str };
        };

        const measureTextWidth = (fontKey, fontStr, text) => {
            const cacheKey = `${fontKey}|${text}`;
            if (measureCache.has(cacheKey)) return measureCache.get(cacheKey);
            ctx.font = fontStr;
            const width = ctx.measureText(text).width;
            measureCache.set(cacheKey, width);
            return width;
        };

        sheet.rows.forEach(row => {
            if (!row || !row.cells?.length) return;
            row.cells.forEach((cell, c) => {
                if (!cell || c < startCol || c > endCol) return;
                const value = cell.showVal;
                if (!value) return;

                const font = cell.style?.font || {};
                const { key, str } = getFontInfo(font);
                const lines = String(value).split('\n');
                let maxWidth = maxWidthByCol.get(c) || 0;
                for (const line of lines) {
                    const width = measureTextWidth(key, str, line) + 16;
                    if (width > maxWidth) maxWidth = width;
                }
                if (maxWidth > 0) maxWidthByCol.set(c, maxWidth);
            });
        });

        maxWidthByCol.forEach((maxWidth, c) => {
            const col = sheet.getCol(c);
            const newWidth = Math.min(Math.round(maxWidth), 500);
            if (newWidth === col._width) return;
            results.push({ col, newWidth, oldWidth: col._width });
            col._width = newWidth;
        });

        // 批量 UndoRedo
        if (results.length > 0) {
            sheet.SN.UndoRedo.add({
                undo: () => results.forEach(({ col, oldWidth }) => col._width = oldWidth),
                redo: () => results.forEach(({ col, newWidth }) => col._width = newWidth)
            });
        }
    }
};

function getActiveArea(SN) {
    const areas = SN.activeSheet.activeAreas;
    return areas[areas.length - 1];
}

function getActiveAreas(SN) {
    const areas = SN.activeSheet.activeAreas;
    if (Array.isArray(areas) && areas.length > 0) return areas;
    return [{ s: SN.activeSheet.activeCell, e: SN.activeSheet.activeCell }];
}

function resolveOutlineAxis(sheet, areas) {
    const isFullColumnSelection = areas.every(area =>
        area?.s?.r === 0 &&
        area?.e?.r === sheet.rowCount - 1
    );
    const hasColumnSpan = areas.some(area =>
        area?.s?.c !== 0 || area?.e?.c !== sheet.colCount - 1
    );
    return isFullColumnSelection && hasColumnSpan ? 'col' : 'row';
}

// ============================================================
// Context Menu Adapters - 右键菜单
// ============================================================

export function hidRows() {
    const area = getActiveArea(this.SN);
    const sheet = this.SN.activeSheet;
    Core.hideRows(sheet, area.s.r, area.e.r);
    this.SN._recordChange({ type: 'row', action: 'hide', sheet, startRow: area.s.r, endRow: area.e.r });
    this.SN.containerDom.querySelector('.sn-row-rc').style.display = 'none';
    this.SN._r();
}

export function delRows() {
    const area = getActiveArea(this.SN);
    Core.deleteRows(this.SN.activeSheet, area.s.r, area.e.r - area.s.r + 1);
    this.SN.containerDom.querySelector('.sn-row-rc').style.display = 'none';
    this.SN._r();
}

export function addRows(event, type) {
    if (event.target.tagName == 'INPUT') return;
    const area = getActiveArea(this.SN);
    const start = type == 'top' ? area.s.r : area.e.r + 1;
    const count = Number(event.target.querySelector('input').value);
    Core.insertRows(this.SN.activeSheet, start, count);
    this.SN.containerDom.querySelector('.sn-row-rc').style.display = 'none';
    this.SN._r();
}

export function delCols() {
    const area = getActiveArea(this.SN);
    Core.deleteCols(this.SN.activeSheet, area.s.c, area.e.c - area.s.c + 1);
    this.SN.containerDom.querySelector('.sn-col-rc').style.display = 'none';
    this.SN._r();
}

export function hidCols() {
    const area = getActiveArea(this.SN);
    const sheet = this.SN.activeSheet;
    Core.hideCols(sheet, area.s.c, area.e.c);
    this.SN._recordChange({ type: 'col', action: 'hide', sheet, startCol: area.s.c, endCol: area.e.c });
    this.SN.containerDom.querySelector('.sn-col-rc').style.display = 'none';
    this.SN._r();
}

export function addCols(event, type) {
    if (event.target.tagName == 'INPUT') return;
    const area = getActiveArea(this.SN);
    const start = type == 'left' ? area.s.c : area.e.c + 1;
    const count = Number(event.target.querySelector('input').value);
    Core.insertCols(this.SN.activeSheet, start, count);
    this.SN.containerDom.querySelector('.sn-col-rc').style.display = 'none';
    this.SN._r();
}

// ============================================================
// Toolbar Adapters - 工具栏（弹窗交互 + UndoRedo 封装）
// ============================================================

function closeDropdownByEvent(event) {
    const dropdown = event?.target?.closest('.sn-dropdown');
    if (!dropdown) return;
    dropdown.classList.add('closing');
    setTimeout(() => {
        dropdown.classList.remove('active', 'closing');
    }, 200);
}

export function applyRowColInsertByPos(event, pos) {
    event?.stopPropagation?.();
    const item = event?.target?.closest('.sn-rowcol-insert-item');
    if (!item) return;
    const countInput = item.querySelector('[data-field="rowcol-insert-count"]');
    if (!countInput) return;
    const isCol = pos === 'left' || pos === 'right';

    const area = getActiveArea(this.SN);
    const sheet = this.SN.activeSheet;
    let count = Math.round(Number(countInput.value));
    if (!Number.isFinite(count) || count < 1) {
        const rowCount = area.e.r - area.s.r + 1;
        const colCount = area.e.c - area.s.c + 1;
        count = isCol ? colCount : rowCount;
    }
    count = Math.min(Math.max(count, 1), 10000);

    if (!isCol) {
        const start = pos === 'above' ? area.s.r : area.e.r + 1;
        Core.insertRows(sheet, start, count);
    } else {
        const start = pos === 'left' ? area.s.c : area.e.c + 1;
        Core.insertCols(sheet, start, count);
    }

    this.SN._r();
    closeDropdownByEvent(event);
}

// 插入行弹窗（RowColHandler 自带 UndoRedo，无需封装）
export async function insertRowDialog() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const area = getActiveArea(SN);
    const t = (key, params = {}) => SN.t(key, params);

    try {
        const result = await SN.Utils.modal({
            titleKey: 'action.rowCol.modal.m001.title',
            title: t('action.rowCol.modal.m001.title'),
            content: `
                <div style="display:flex;flex-direction:column;gap:12px;min-width:240px;">
                    <div>
                        <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.insert.countLabel')}</label>
                        <input type="number" data-field="insert-row-count" value="${area.e.r - area.s.r + 1}" min="1" max="1000"
                            style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
                    </div>
                    <div>
                        <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.insert.positionLabel')}</label>
                        <select data-field="insert-row-pos" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;">
                            <option value="above">${t('action.rowCol.content.insertRow.positionAbove')}</option>
                            <option value="below">${t('action.rowCol.content.insertRow.positionBelow')}</option>
                        </select>
                    </div>
                </div>
            `,
            confirmKey: 'action.rowCol.modal.m001.confirm',
            confirmText: t('action.rowCol.modal.m001.confirm'),
            cancelKey: 'action.rowCol.modal.m001.cancel',
            cancelText: t('action.rowCol.modal.m001.cancel')
        });

        const count = parseInt(result.bodyEl.querySelector('[data-field="insert-row-count"]').value) || 1;
        const pos = result.bodyEl.querySelector('[data-field="insert-row-pos"]').value;
        Core.insertRows(sheet, pos === 'above' ? area.s.r : area.e.r + 1, count);
        SN._r();
    } catch (e) { /* canceled */ }
}

// 插入列弹窗
export async function insertColDialog() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const area = getActiveArea(SN);
    const t = (key, params = {}) => SN.t(key, params);

    try {
        const result = await SN.Utils.modal({
            titleKey: 'action.rowCol.modal.m002.title',
            title: t('action.rowCol.modal.m002.title'),
            content: `
                <div style="display:flex;flex-direction:column;gap:12px;min-width:240px;">
                    <div>
                        <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.insert.countLabel')}</label>
                        <input type="number" data-field="insert-col-count" value="${area.e.c - area.s.c + 1}" min="1" max="100"
                            style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
                    </div>
                    <div>
                        <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.insert.positionLabel')}</label>
                        <select data-field="insert-col-pos" style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;">
                            <option value="left">${t('action.rowCol.content.insertCol.positionLeft')}</option>
                            <option value="right">${t('action.rowCol.content.insertCol.positionRight')}</option>
                        </select>
                    </div>
                </div>
            `,
            confirmKey: 'action.rowCol.modal.m002.confirm',
            confirmText: t('action.rowCol.modal.m002.confirm'),
            cancelKey: 'action.rowCol.modal.m002.cancel',
            cancelText: t('action.rowCol.modal.m002.cancel')
        });

        const count = parseInt(result.bodyEl.querySelector('[data-field="insert-col-count"]').value) || 1;
        const pos = result.bodyEl.querySelector('[data-field="insert-col-pos"]').value;
        Core.insertCols(sheet, pos === 'left' ? area.s.c : area.e.c + 1, count);
        SN._r();
    } catch (e) { /* canceled */ }
}

// 删除行（RowColHandler 自带 UndoRedo）
export function deleteRowAction() {
    const area = getActiveArea(this.SN);
    Core.deleteRows(this.SN.activeSheet, area.s.r, area.e.r - area.s.r + 1);
    this.SN._r();
}

// 删除列
export function deleteColAction() {
    const area = getActiveArea(this.SN);
    Core.deleteCols(this.SN.activeSheet, area.s.c, area.e.c - area.s.c + 1);
    this.SN._r();
}

// 隐藏行
export function hideRowAction() {
    const sheet = this.SN.activeSheet;
    const area = getActiveArea(this.SN);
    Core.hideRows(sheet, area.s.r, area.e.r);
    this.SN._r();
}

// 隐藏列
export function hideColAction() {
    const sheet = this.SN.activeSheet;
    const area = getActiveArea(this.SN);
    Core.hideCols(sheet, area.s.c, area.e.c);
    this.SN._r();
}

// 取消隐藏行
export function unhideRowAction() {
    const sheet = this.SN.activeSheet;
    const area = getActiveArea(this.SN);
    Core.unhideRows(sheet, area.s.r, area.e.r);
    this.SN._r();
}

// 取消隐藏列
export function unhideColAction() {
    const sheet = this.SN.activeSheet;
    const area = getActiveArea(this.SN);
    Core.unhideCols(sheet, area.s.c, area.e.c);
    this.SN._r();
}

// 取消隐藏所有行
export function unhideAllRowsAction() {
    const sheet = this.SN.activeSheet;
    Core.unhideAllRows(sheet);
    this.SN._r();
}

// 取消隐藏所有列
export function unhideAllColsAction() {
    const sheet = this.SN.activeSheet;
    Core.unhideAllCols(sheet);
    this.SN._r();
}

// 创建组（行）
export function groupRows() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const areas = getActiveAreas(SN);
    const Utils = SN.Utils;
    const axis = resolveOutlineAxis(sheet, areas);
    const result = axis === 'col' ? sheet.groupCols(areas) : sheet.groupRows(areas);
    if (result?.canceled) return;
    const changedCount = result?.changedCount || 0;
    const targetCount = axis === 'col' ? (result?.colCount || 0) : (result?.rowCount || 0);
    const limitedCount = result?.limitedCount || 0;
    const unitLabel = axis === 'col' ? SN.t('action.rowCol.unit.col') : SN.t('action.rowCol.unit.row');

    if (targetCount === 0) {
        Utils.toast(SN.t('action.rowCol.toast.group.selectTarget', { unitLabel }));
        return;
    }
    if (changedCount === 0) {
        Utils.toast(SN.t('action.rowCol.toast.group.maxLevelReached', { unitLabel }));
        return;
    }

    SN._r();
    if (limitedCount > 0) {
        Utils.toast(SN.t('action.rowCol.toast.group.completedWithLimited', { changedCount, unitLabel, limitedCount }));
    } else {
        Utils.toast(SN.t('action.rowCol.toast.group.completed', { changedCount, unitLabel }));
    }
}

// 取消组合（行）
export function ungroupRows() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const areas = getActiveAreas(SN);
    const Utils = SN.Utils;
    const axis = resolveOutlineAxis(sheet, areas);
    const result = axis === 'col' ? sheet.ungroupCols(areas) : sheet.ungroupRows(areas);
    if (result?.canceled) return;
    const changedCount = result?.changedCount || 0;
    const targetCount = axis === 'col' ? (result?.colCount || 0) : (result?.rowCount || 0);
    const unitLabel = axis === 'col' ? SN.t('action.rowCol.unit.col') : SN.t('action.rowCol.unit.row');

    if (targetCount === 0) {
        Utils.toast(SN.t('action.rowCol.toast.ungroup.selectTarget', { unitLabel }));
        return;
    }
    if (changedCount === 0) {
        Utils.toast(SN.t('action.rowCol.toast.ungroup.noGroup', { unitLabel }));
        return;
    }

    SN._r();
    Utils.toast(SN.t('action.rowCol.toast.ungroup.completed', { changedCount, unitLabel }));
}

// 设置行高弹窗
export async function setRowHeightDialog() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const area = getActiveArea(SN);
    const currentHeight = sheet.getRow(area.s.r).height;
    const t = (key, params = {}) => SN.t(key, params);

    try {
        const result = await SN.Utils.modal({
            titleKey: 'action.rowCol.modal.m003.title',
            title: t('action.rowCol.modal.m003.title'),
            maxWidth: '400px',
            content: `
                <div style="min-width:200px;">
                    <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.rowHeight.label')}</label>
                    <input type="number" data-field="row-height" value="${currentHeight}" min="1" max="500"
                        style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
                </div>
            `,
            confirmKey: 'action.rowCol.modal.m003.confirm',
            confirmText: t('action.rowCol.modal.m003.confirm'),
            cancelKey: 'action.rowCol.modal.m003.cancel',
            cancelText: t('action.rowCol.modal.m003.cancel')
        });

        const height = parseInt(result.bodyEl.querySelector('[data-field="row-height"]').value) || currentHeight;
        Core.setRowsHeight(sheet, area.s.r, area.e.r, height);
        SN._r();
    } catch (e) { /* canceled */ }
}

// 设置列宽弹窗
export async function setColWidthDialog() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const area = getActiveArea(SN);
    const currentWidth = sheet.getCol(area.s.c).width;
    const t = (key, params = {}) => SN.t(key, params);

    try {
        const result = await SN.Utils.modal({
            titleKey: 'action.rowCol.modal.m004.title',
            title: t('action.rowCol.modal.m004.title'),
            maxWidth: '400px',
            content: `
                <div style="min-width:200px;">
                    <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.colWidth.label')}</label>
                    <input type="number" data-field="col-width" value="${currentWidth}" min="1" max="1000"
                        style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
                </div>
            `,
            confirmKey: 'action.rowCol.modal.m004.confirm',
            confirmText: t('action.rowCol.modal.m004.confirm'),
            cancelKey: 'action.rowCol.modal.m004.cancel',
            cancelText: t('action.rowCol.modal.m004.cancel')
        });

        const width = parseInt(result.bodyEl.querySelector('[data-field="col-width"]').value) || currentWidth;
        Core.setColsWidth(sheet, area.s.c, area.e.c, width);
        SN._r();
    } catch (e) { /* canceled */ }
}

// 设置默认行高弹窗
export async function setDefaultRowHeightDialog() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const currentHeight = sheet.defaultRowHeight;
    const t = (key, params = {}) => SN.t(key, params);

    try {
        const result = await SN.Utils.modal({
            titleKey: 'action.rowCol.modal.m005.title',
            title: t('action.rowCol.modal.m005.title'),
            maxWidth: '400px',
            content: `
                <div style="min-width:200px;">
                    <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.defaultRowHeight.label')}</label>
                    <input type="number" data-field="default-row-height" value="${currentHeight}" min="1" max="500"
                        style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
                </div>
            `,
            confirmKey: 'action.rowCol.modal.m005.confirm',
            confirmText: t('action.rowCol.modal.m005.confirm'),
            cancelKey: 'action.rowCol.modal.m005.cancel',
            cancelText: t('action.rowCol.modal.m005.cancel')
        });

        const height = parseInt(result.bodyEl.querySelector('[data-field="default-row-height"]').value) || currentHeight;
        if (height === currentHeight) return;
        sheet.defaultRowHeight = height;
        SN._r();
    } catch (e) { /* canceled */ }
}

// 设置默认列宽弹窗
export async function setDefaultColWidthDialog() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const currentWidth = sheet.defaultColWidth;
    const t = (key, params = {}) => SN.t(key, params);

    try {
        const result = await SN.Utils.modal({
            titleKey: 'action.rowCol.modal.m006.title',
            title: t('action.rowCol.modal.m006.title'),
            maxWidth: '400px',
            content: `
                <div style="min-width:200px;">
                    <label style="display:block;margin-bottom:4px;font-size:13px;">${t('action.rowCol.content.defaultColWidth.label')}</label>
                    <input type="number" data-field="default-col-width" value="${currentWidth}" min="1" max="1000"
                        style="width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:4px;box-sizing:border-box;">
                </div>
            `,
            confirmKey: 'action.rowCol.modal.m006.confirm',
            confirmText: t('action.rowCol.modal.m006.confirm'),
            cancelKey: 'action.rowCol.modal.m006.cancel',
            cancelText: t('action.rowCol.modal.m006.cancel')
        });

        const width = parseInt(result.bodyEl.querySelector('[data-field="default-col-width"]').value) || currentWidth;
        if (width === currentWidth) return;
        sheet.defaultColWidth = width;
        SN._r();
    } catch (e) { /* canceled */ }
}

// 自适应行高
export function autoFitRowHeight() {
    const sheet = this.SN.activeSheet;
    const area = getActiveArea(this.SN);
    Core.autoFitRowsHeight(sheet, area.s.r, area.e.r);
    this.SN._r();
}

// 自适应列宽
export function autoFitColWidth() {
    const sheet = this.SN.activeSheet;
    const area = getActiveArea(this.SN);
    Core.autoFitColsWidth(sheet, area.s.c, area.e.c);
    this.SN._r();
}
