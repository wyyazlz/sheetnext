/**
 * 筛选操作模块
 * 提供自动筛选相关的工具栏操作方法
 */

import SortDialog from '../components/SortDialog.js';

// 排序对话框实例缓存
let sortDialogInstance = null;

function getActiveFilterScope(sheet) {
    const table = sheet.Table?.getTableAt?.(sheet.activeCell.r, sheet.activeCell.c);
    if (table?.autoFilterEnabled) {
        const scopeId = sheet.AutoFilter.getTableScopeId(table.id);
        if (sheet.AutoFilter.isScopeEnabled(scopeId)) {
            return { scopeId, table };
        }
    }
    return { scopeId: null, table: null };
}

/**
 * 切换自动筛选
 * 如果没有启用筛选，则启用；如果已启用，则关闭
 */
export function toggleAutoFilter() {
    const sheet = this.SN.activeSheet;
    const autoFilter = sheet.AutoFilter;
    const { scopeId, table } = getActiveFilterScope(sheet);

    if (table) {
        table.autoFilterEnabled = !autoFilter.isScopeEnabled(scopeId);
        return;
    }

    if (autoFilter.enabled) {
        // 已启用，关闭筛选
        autoFilter.clear();
    } else {
        // 未启用，设置筛选范围
        autoFilter.setRange();
    }
}

/**
 * 清除所有筛选条件（保留筛选范围）
 */
export function clearFilter() {
    const sheet = this.SN.activeSheet;
    const autoFilter = sheet.AutoFilter;
    const { scopeId } = getActiveFilterScope(sheet);

    if (scopeId && autoFilter.isScopeEnabled(scopeId)) {
        autoFilter.clearAllFilters(scopeId);
    } else if (autoFilter.enabled) {
        autoFilter.clearAllFilters();
    } else {
        this.SN.Utils.toast(this.SN.t('action.filter.toast.msg001'));
    }
}

/**
 * 重新应用筛选
 */
export function reapplyFilter() {
    const sheet = this.SN.activeSheet;
    const autoFilter = sheet.AutoFilter;
    const { scopeId } = getActiveFilterScope(sheet);

    if (scopeId && autoFilter.isScopeEnabled(scopeId)) {
        autoFilter._applyFilters(scopeId);
        this.SN.Utils.toast(this.SN.t('action.filter.toast.msg002'));
    } else if (autoFilter.enabled) {
        autoFilter._applyFilters();
        this.SN.Utils.toast(this.SN.t('action.filter.toast.msg002'));
    } else {
        this.SN.Utils.toast(this.SN.t('action.filter.toast.msg001'));
    }
}

/**
 * 升序排序
 */
export function sortAsc() {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;

    if (areas.length === 0) {
        return this.SN.Utils.toast(this.SN.t('action.filter.toast.msg003'));
    }

    const area = areas[areas.length - 1];
    const sortCol = sheet.activeCell.c;

    if (sheet.rangeSort([{ col: sortCol, order: 'asc' }], area)) {
        this.SN._r();
    }
}

/**
 * 降序排序
 */
export function sortDesc() {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;

    if (areas.length === 0) {
        return this.SN.Utils.toast(this.SN.t('action.filter.toast.msg003'));
    }

    const area = areas[areas.length - 1];
    const sortCol = sheet.activeCell.c;

    if (sheet.rangeSort([{ col: sortCol, order: 'desc' }], area)) {
        this.SN._r();
    }
}

/**
 * 自定义排序对话框
 */
export function customSort() {
    if (!sortDialogInstance) {
        sortDialogInstance = new SortDialog(this.SN);
    }
    sortDialogInstance.show();
}
