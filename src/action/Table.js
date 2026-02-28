/**
 * 超级表操作模块
 * 负责创建、管理表格的用户交互
 */
import { TABLE_STYLE_GROUPS, getStylePreviewColors } from '../core/Table/TableStyle.js';
import { createTableSlicers } from '../core/Slicer/index.js';

/**
 * 打开创建表格对话框
 */
export function createTable() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    // 获取当前选区
    const area = sheet.activeAreas[0];
    if (!area) {
        Utils.toast(SN.t('action.table.toast.pleaseSelectARangeFirst'));
        return;
    }

    // 格式化区域引用（绝对路径带$）
    const refStr = Utils.rangeNumToStr(area, true);

    const bodyHTML = `
        <div class="sn-table-dialog">
            <div class="sn-form-group">
                <label>${t('action.table.content.sourceRangeLabel')}</label>
                <input type="text" class="sn-table-ref" value="${refStr}" placeholder="${t('action.table.content.sourceRangePlaceholder')}">
            </div>
            <div class="sn-form-group">
                <label>
                    <input type="checkbox" class="sn-table-has-header" checked>
                    ${t('action.table.content.hasHeader')}
                </label>
            </div>
        </div>
    `;

    Utils.modal({
        titleKey: 'action.table.modal.m001.title', title: t('action.table.modal.m001.title'),
        body: bodyHTML,
        maxWidth: '360px',
        confirmKey: 'action.table.modal.m001.confirm', confirmText: t('action.table.modal.m001.confirm'),
        cancelKey: 'action.table.modal.m001.cancel', cancelText: t('action.table.modal.m001.cancel'),
        onConfirm: (bodyEl) => {
            const refInput = bodyEl.querySelector('.sn-table-ref');
            const hasHeaderCheckbox = bodyEl.querySelector('.sn-table-has-header');

            const ref = refInput.value.trim();
            if (!ref) {
                Utils.toast(SN.t('action.table.toast.pleaseEnterAValidRange'));
                return false;
            }

            try {
                const range = sheet.rangeStrToNum(ref);

                // 检查是否与现有表格重叠
                for (const table of sheet.Table.getAll()) {
                    const tr = table.range;
                    if (tr && !(range.e.r < tr.s.r || range.s.r > tr.e.r ||
                               range.e.c < tr.s.c || range.s.c > tr.e.c)) {
                        Utils.toast(SN.t('action.table.toast.theSelectedRangeOverlapsAn'));
                        return false;
                    }
                }

                // 创建表格（使用默认样式）
                const newTable = sheet.Table.createFromSelection(range, {
                    styleId: 'TableStyleMedium2',
                    showHeaderRow: hasHeaderCheckbox.checked,
                    showRowStripes: true,
                    autoFilterEnabled: true
                });

                // 刷新画布
                this.SN.Canvas.r();

                // 创建成功后自动切换到"表设计"选项卡
                this.SN.Layout?.updateContextualToolbar({
                    type: 'table',
                    data: newTable,
                    autoSwitch: true
                });
                return true;
            } catch (e) {
                Utils.toast(e.message || SN.t('action.table.toast.createFailed'));
                return false;
            }
        }
    });
}

/**
 * 套用表格格式 - 从当前选区创建表格并应用指定样式
 * @param {string} styleId - 表格样式ID
 */
export function formatAsTable(styleId) {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    // 如果当前已在表格中，直接改样式
    const { r, c } = sheet.activeCell;
    const existingTable = sheet.Table.getTableAt(r, c);
    if (existingTable) {
        existingTable.styleId = styleId;
        this.SN.Canvas.r();
        return;
    }

    const area = sheet.activeAreas[0];
    if (!area) {
        Utils.toast(SN.t('action.table.toast.pleaseSelectARangeFirst'));
        return;
    }

    const refStr = Utils.rangeNumToStr(area, true);

    const bodyHTML = `
        <div class="sn-table-dialog">
            <div class="sn-form-group">
                <label>${t('action.table.content.sourceRangeLabel')}</label>
                <input type="text" class="sn-table-ref" value="${refStr}" placeholder="${t('action.table.content.sourceRangePlaceholder')}">
            </div>
            <div class="sn-form-group">
                <label>
                    <input type="checkbox" class="sn-table-has-header" checked>
                    ${t('action.table.content.hasHeader')}
                </label>
            </div>
        </div>
    `;

    Utils.modal({
        titleKey: 'action.table.modal.m002.title', title: t('action.table.modal.m002.title'),
        body: bodyHTML,
        maxWidth: '360px',
        confirmKey: 'action.table.modal.m002.confirm', confirmText: t('action.table.modal.m002.confirm'),
        cancelKey: 'action.table.modal.m002.cancel', cancelText: t('action.table.modal.m002.cancel'),
        onConfirm: (bodyEl) => {
            const refInput = bodyEl.querySelector('.sn-table-ref');
            const hasHeaderCheckbox = bodyEl.querySelector('.sn-table-has-header');
            const ref = refInput.value.trim();
            if (!ref) {
                Utils.toast(SN.t('action.table.toast.pleaseEnterAValidRange'));
                return false;
            }

            try {
                const range = sheet.rangeStrToNum(ref);

                for (const table of sheet.Table.getAll()) {
                    const tr = table.range;
                    if (tr && !(range.e.r < tr.s.r || range.s.r > tr.e.r ||
                               range.e.c < tr.s.c || range.s.c > tr.e.c)) {
                        Utils.toast(SN.t('action.table.toast.theSelectedRangeOverlapsAn'));
                        return false;
                    }
                }

                const newTable = sheet.Table.createFromSelection(range, {
                    styleId: styleId,
                    showHeaderRow: hasHeaderCheckbox.checked,
                    showRowStripes: true,
                    autoFilterEnabled: true
                });

                this.SN.Canvas.r();
                this.SN.Layout?.updateContextualToolbar({
                    type: 'table',
                    data: newTable,
                    autoSwitch: true
                });
                return true;
            } catch (e) {
                Utils.toast(e.message || SN.t('action.table.toast.createFailed'));
                return false;
            }
        }
    });
}

/**
 * 转换表为普通区域
 */
export function convertTableToRange() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    Utils.modal({
        titleKey: 'action.table.modal.m003.title',
        title: t('action.table.modal.m003.title'),
        body: `<p>${t('action.table.content.convertToRangeBody')}</p>`,
        confirmKey: 'action.table.modal.m003.confirm',
        confirmText: t('action.table.modal.m003.confirm'),
        cancelKey: 'action.table.modal.m003.cancel',
        cancelText: t('action.table.modal.m003.cancel'),
        onConfirm: () => {
            table.convertToRange();
            this.SN.Canvas.r();
            this.SN.Layout?.updateContextualToolbar(null);
            Utils.toast(SN.t('action.table.toast.msg005'));
            return true;
        }
    });
}

/**
 * 调整表格大小
 */
export function resizeTable() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    const currentRef = table.ref;

    Utils.modal({
        titleKey: 'action.table.modal.m004.title',
        title: t('action.table.modal.m004.title'),
        maxWidth: '400px',
        body: `
            <div class="sn-form-group">
                <label>${t('action.table.content.resizeRangeLabel')}</label>
                <input type="text" class="sn-table-resize-ref" value="${currentRef}" placeholder="${t('action.table.content.resizeRangePlaceholder')}">
            </div>
        `,
        confirmKey: 'action.table.modal.m004.confirm',
        confirmText: t('action.table.modal.m004.confirm'),
        cancelKey: 'action.table.modal.m004.cancel',
        cancelText: t('action.table.modal.m004.cancel'),
        onConfirm: (bodyEl) => {
            const refInput = bodyEl.querySelector('.sn-table-resize-ref');
            const newRef = refInput.value.trim();

            if (!newRef) {
                Utils.toast(SN.t('action.table.toast.msg006'));
                return false;
            }

            try {
                table.resize(newRef);
                this.SN.Canvas.r();
                Utils.toast(SN.t('action.table.toast.msg007'));
                return true;
            } catch (e) {
                Utils.toast(e.message || SN.t('action.table.toast.resizeFailed'));
                return false;
            }
        }
    });
}

/**
 * 更改表格样式（从工具栏菜单调用）
 * @param {string} styleId - 样式ID
 */
export function changeTableStyle(styleId) {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    const resolveCurrentTable = () => {
        const { r, c } = sheet.activeCell || {};
        const tableAtActiveCell = Number.isInteger(r) && Number.isInteger(c)
            ? sheet.Table.getTableAt(r, c)
            : null;
        if (tableAtActiveCell) return tableAtActiveCell;

        const layout = this.SN.Layout;
        if (layout?._currentContextType !== 'table') return null;
        const contextTable = layout._currentContextData;
        if (!contextTable || contextTable.sheet !== sheet) return null;

        return sheet.Table.get(contextTable.id) || contextTable;
    };

    const table = resolveCurrentTable();

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    if (styleId) {
        table.styleId = styleId;
        this.SN.Canvas.r();
        return;
    }

    const buildAllStyles = () => {
        const allStyles = [...TABLE_STYLE_GROUPS.light, ...TABLE_STYLE_GROUPS.medium, ...TABLE_STYLE_GROUPS.dark];
        return allStyles.map(sid => {
            const colors = getStylePreviewColors(sid, this.SN);
            const selected = sid === table.styleId ? 'selected' : '';
            return `
                <div class="sn-table-style-item ${selected}" data-style="${sid}" title="${sid}">
                    <div class="sn-table-style-preview">
                        <div class="sn-table-style-row header" style="background:${colors.header};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row" style="background:${colors.stripe1};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row stripe" style="background:${colors.stripe2};border-color:${colors.border}"></div>
                    </div>
                </div>
            `;
        }).join('');
    };

    let selectedStyle = table.styleId;

    Utils.modal({
        titleKey: 'action.table.modal.m005.title',
        title: t('action.table.modal.m005.title'),
        body: `
            <div class="sn-table-styles-grid">
                ${buildAllStyles()}
            </div>
        `,
        confirmKey: 'action.table.modal.m005.confirm',
        confirmText: t('action.table.modal.m005.confirm'),
        cancelKey: 'action.table.modal.m005.cancel',
        cancelText: t('action.table.modal.m005.cancel'),
        bodyClass: 'sn-table-style-picker-body',
        onConfirm: () => {
            table.styleId = selectedStyle;
            this.SN.Canvas.r();
            return true;
        },
        onOpen: (bodyEl) => {
            const styleItems = bodyEl.querySelectorAll('.sn-table-style-item');
            styleItems.forEach(item => {
                item.addEventListener('click', () => {
                    styleItems.forEach(i => i.classList.remove('selected'));
                    item.classList.add('selected');
                    selectedStyle = item.dataset.style;
                });
            });
        }
    });
}

/**
 * 切换表格选项
 * @param {string} option - 选项名称
 */
export function toggleTableOption(option) {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    switch (option) {
        case 'headerRow':
            table.showHeaderRow = !table.showHeaderRow;
            break;
        case 'totalsRow':
            table.showTotalsRow = !table.showTotalsRow;
            break;
        case 'firstColumn':
            table.showFirstColumn = !table.showFirstColumn;
            break;
        case 'lastColumn':
            table.showLastColumn = !table.showLastColumn;
            break;
        case 'rowStripes':
            table.showRowStripes = !table.showRowStripes;
            break;
        case 'columnStripes':
            table.showColumnStripes = !table.showColumnStripes;
            break;
        case 'autoFilter':
            table.autoFilterEnabled = !table.autoFilterEnabled;
            break;
    }

    this.SN.Canvas.r();
}

/**
 * 重命名表格
 * @param {string} newName - 新表名（可选，不传则弹出对话框）
 */
export function renameTable(newName) {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    if (newName) {
        try {
            table.name = newName;
            this.SN.Canvas.r();
        } catch (e) {
            Utils.toast(e.message || SN.t('action.table.toast.renameFailed'));
        }
        return;
    }

    Utils.modal({
        titleKey: 'action.table.modal.m006.title',
        title: t('action.table.modal.m006.title'),
        body: `
            <div class="sn-form-group">
                <label>${t('action.table.content.renameLabel')}</label>
                <input type="text" class="sn-table-name-input" value="${table.displayName || table.name}" placeholder="${t('action.table.content.renamePlaceholder')}">
            </div>
        `,
        confirmKey: 'action.table.modal.m006.confirm',
        confirmText: t('action.table.modal.m006.confirm'),
        cancelKey: 'action.table.modal.m006.cancel',
        cancelText: t('action.table.modal.m006.cancel'),
        onConfirm: (bodyEl) => {
            const nameInput = bodyEl.querySelector('.sn-table-name-input');
            const name = nameInput.value.trim();

            if (!name) {
                Utils.toast(SN.t('action.table.toast.msg008'));
                return false;
            }

            try {
                table.name = name;
                this.SN.Canvas.r();
                Utils.toast(SN.t('action.table.toast.msg009'));
                return true;
            } catch (e) {
                Utils.toast(e.message || SN.t('action.table.toast.renameFailed'));
                return false;
            }
        }
    });
}

/**
 * 删除表格（保留数据，只移除表格定义）
 */
export function deleteTable() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    Utils.modal({
        titleKey: 'action.table.modal.m007.title',
        title: t('action.table.modal.m007.title'),
        body: `<p>${t('action.table.content.deleteTableBody')}</p>`,
        confirmKey: 'action.table.modal.m007.confirm',
        confirmText: t('action.table.modal.m007.confirm'),
        cancelKey: 'action.table.modal.m007.cancel',
        cancelText: t('action.table.modal.m007.cancel'),
        onConfirm: () => {
            table.convertToRange();
            this.SN.Canvas.r();
            this.SN.Layout?.updateContextualToolbar(null);
            Utils.toast(SN.t('action.table.toast.msg010'));
            return true;
        }
    });
}

/**
 * 清除表格样式（设为无样式）
 */
export function clearTableStyle() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    table.styleId = null;
    this.SN.Canvas.r();
}

/**
 * 插入汇总行中的函数
 * @param {string} funcType - 函数类型: sum, average, count, countNums, max, min, stdDev, var
 */
export function insertTableTotalsFunction(funcType) {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg004'));
        return;
    }

    // Get current column index in table
    const colIndex = c - table.range.s.c;

    try {
        table.setTotalsFunction(colIndex, funcType);
        this.SN.Canvas.r();
    } catch (e) {
        Utils.toast(e.message || SN.t('action.table.toast.setTotalsFunctionFailed'));
    }
}

/**
 * 删除重复项
 * 复刻Excel删除重复项功能
 */
export function removeDuplicates() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    const area = sheet.activeAreas[0];
    if (!area) {
        Utils.toast(SN.t('action.table.toast.msg011'));
        return;
    }

    const startCol = area.s.c;
    const endCol = area.e.c;
    const startRow = area.s.r;
    const columns = [];

    for (let c = startCol; c <= endCol; c++) {
        const headerCell = sheet.getCell(startRow, c);
        const headerValue = headerCell.v ?? Utils.numToChar(c);
        columns.push({
            index: c - startCol,
            col: c,
            name: String(headerValue)
        });
    }

    const columnsHTML = columns.map((col) => `
        <div class="sn-duplicate-col-item">
            <label>
                <input type="checkbox" class="sn-duplicate-col-check" data-col="${col.col}" checked>
                <span>${col.name}</span>
            </label>
        </div>
    `).join('');

    const bodyHTML = `
        <div class="sn-remove-duplicates-dialog">
            <div class="sn-form-group">
                <label>
                    <input type="checkbox" class="sn-has-header" checked>
                    ${t('action.table.content.removeDuplicates.hasHeader')}
                </label>
            </div>
            <div class="sn-form-group">
                <label>${t('action.table.content.removeDuplicates.selectColumns')}</label>
                <div class="sn-duplicate-actions">
                    <button type="button" class="sn-btn-link sn-select-all">${t('action.table.content.removeDuplicates.selectAll')}</button>
                    <button type="button" class="sn-btn-link sn-deselect-all">${t('action.table.content.removeDuplicates.deselectAll')}</button>
                </div>
                <div class="sn-duplicate-columns">
                    ${columnsHTML}
                </div>
            </div>
        </div>
    `;

    Utils.modal({
        titleKey: 'action.table.modal.m008.title',
        title: t('action.table.modal.m008.title'),
        body: bodyHTML,
        maxWidth: '360px',
        confirmKey: 'action.table.modal.m008.confirm',
        confirmText: t('action.table.modal.m008.confirm'),
        cancelKey: 'action.table.modal.m008.cancel',
        cancelText: t('action.table.modal.m008.cancel'),
        onOpen: (bodyEl) => {
            const selectAll = bodyEl.querySelector('.sn-select-all');
            const deselectAll = bodyEl.querySelector('.sn-deselect-all');
            const checkboxes = bodyEl.querySelectorAll('.sn-duplicate-col-check');

            selectAll?.addEventListener('click', () => {
                checkboxes.forEach(cb => cb.checked = true);
            });
            deselectAll?.addEventListener('click', () => {
                checkboxes.forEach(cb => cb.checked = false);
            });
        },
        onConfirm: (bodyEl) => {
            const hasHeader = bodyEl.querySelector('.sn-has-header').checked;
            const checkboxes = bodyEl.querySelectorAll('.sn-duplicate-col-check:checked');

            if (checkboxes.length === 0) {
                Utils.toast(SN.t('action.table.toast.msg012'));
                return false;
            }

            const selectedCols = Array.from(checkboxes).map(cb => parseInt(cb.dataset.col));
            const dataStartRow = hasHeader ? area.s.r + 1 : area.s.r;
            const dataEndRow = area.e.r;

            const rows = [];
            for (let r = dataStartRow; r <= dataEndRow; r++) {
                const key = selectedCols.map(c => {
                    const cell = sheet.getCell(r, c);
                    return cell.v ?? '';
                }).join('\\u0000');
                rows.push({ row: r, key });
            }

            const seen = new Set();
            const duplicateRows = [];
            for (const item of rows) {
                if (seen.has(item.key)) {
                    duplicateRows.push(item.row);
                } else {
                    seen.add(item.key);
                }
            }

            if (duplicateRows.length === 0) {
                Utils.toast(SN.t('action.table.toast.msg013'));
                return true;
            }

            duplicateRows.sort((a, b) => b - a);
            for (const row of duplicateRows) {
                sheet.deleteRow(row, 1);
            }

            this.SN.Canvas.r();
            Utils.toast(SN.t('action.table.toast.removeDuplicatesDone', {
                duplicateCount: duplicateRows.length,
                uniqueCount: rows.length - duplicateRows.length
            }));
            return true;
        }
    });
}

/**
 * 插入切片器
 */
export function insertSlicer() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const t = (key, params = {}) => SN.t(key, params);

    const { r, c } = sheet.activeCell;
    const table = sheet.Table.getTableAt(r, c);

    if (!table) {
        Utils.toast(SN.t('action.table.toast.msg014'));
        return;
    }

    const columnsHTML = table.columns.map((col, i) => `
        <div class="sn-slicer-col-item">
            <label>
                <input type="checkbox" class="sn-slicer-col-check" data-col="${i}" data-name="${col.name}">
                <span>${col.name}</span>
            </label>
        </div>
    `).join('');

    const bodyHTML = `
        <div class="sn-insert-slicer-dialog">
            <div class="sn-form-group">
                <label>${t('action.table.content.insertSlicer.selectFields')}</label>
                <div class="sn-slicer-columns">
                    ${columnsHTML}
                </div>
            </div>
        </div>
    `;

    Utils.modal({
        titleKey: 'action.table.modal.m009.title',
        title: t('action.table.modal.m009.title'),
        body: bodyHTML,
        maxWidth: '320px',
        confirmKey: 'action.table.modal.m009.confirm',
        confirmText: t('action.table.modal.m009.confirm'),
        cancelKey: 'action.table.modal.m009.cancel',
        cancelText: t('action.table.modal.m009.cancel'),
        onConfirm: (bodyEl) => {
            const checkboxes = bodyEl.querySelectorAll('.sn-slicer-col-check:checked');

            if (checkboxes.length === 0) {
                Utils.toast(SN.t('action.table.toast.msg015'));
                return false;
            }

            const selectedCols = Array.from(checkboxes).map(cb => ({
                index: parseInt(cb.dataset.col, 10),
                name: cb.dataset.name
            }));

            createTableSlicers(this.SN, sheet, table, selectedCols);
            return true;
        }
    });
}
