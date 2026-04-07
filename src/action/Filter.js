/** Filter actions. */

import SortDialog from '../components/SortDialog.js';

let sortDialogInstance = null;

function isCellInRange(cell, range) {
    return !!cell && !!range &&
        cell.r >= range.s.r && cell.r <= range.e.r &&
        cell.c >= range.s.c && cell.c <= range.e.c;
}

function getActiveFilterScope(sheet) {
    const autoFilter = sheet.AutoFilter;
    const table = sheet.Table?.getTableAt?.(sheet.activeCell.r, sheet.activeCell.c);
    if (table?.autoFilterEnabled) {
        const scopeId = autoFilter.getTableScopeId(table.id);
        if (autoFilter.isScopeEnabled(scopeId)) {
            return {
                scopeId,
                table,
                range: autoFilter.getScopeRange(scopeId),
                hasHeader: table.showHeaderRow === true
            };
        }
    }

    const range = autoFilter?.getScopeRange?.();
    if (autoFilter?.enabled && isCellInRange(sheet.activeCell, range)) {
        return {
            scopeId: null,
            table: null,
            range,
            hasHeader: true
        };
    }

    return { scopeId: null, table: null, range: null, hasHeader: false };
}

/** Toggle autofilter. */
export function toggleAutoFilter() {
    const sheet = this.SN.activeSheet;
    const autoFilter = sheet.AutoFilter;
    const { scopeId, table } = getActiveFilterScope(sheet);

    if (table) {
        table.autoFilterEnabled = !autoFilter.isScopeEnabled(scopeId);
        return;
    }

    if (autoFilter.enabled) {
        autoFilter.clear();
    } else {
        autoFilter.setRange();
    }
}

/** Clear filters. */
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

/** Reapply filters. */
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

/** Sort ascending. */
export function sortAsc() {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;
    const { scopeId, range, hasHeader } = getActiveFilterScope(sheet);
    const targetRange = range || (areas.length > 0 ? areas[areas.length - 1] : null);

    if (!targetRange) {
        return this.SN.Utils.toast(this.SN.t('action.filter.toast.msg003'));
    }

    const sortCol = sheet.activeCell.c;
    if (sheet.rangeSort([{ col: sortCol, order: 'asc' }], targetRange, { hasHeader })) {
        if (range) {
            sheet.AutoFilter.setSortState(sortCol, 'asc', scopeId);
            sheet.AutoFilter._applyFilters(scopeId);
        }
        this.SN._r();
    }
}

/** Sort descending. */
export function sortDesc() {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;
    const { scopeId, range, hasHeader } = getActiveFilterScope(sheet);
    const targetRange = range || (areas.length > 0 ? areas[areas.length - 1] : null);

    if (!targetRange) {
        return this.SN.Utils.toast(this.SN.t('action.filter.toast.msg003'));
    }

    const sortCol = sheet.activeCell.c;
    if (sheet.rangeSort([{ col: sortCol, order: 'desc' }], targetRange, { hasHeader })) {
        if (range) {
            sheet.AutoFilter.setSortState(sortCol, 'desc', scopeId);
            sheet.AutoFilter._applyFilters(scopeId);
        }
        this.SN._r();
    }
}

/** Show custom sort dialog. */
export function customSort() {
    if (!sortDialogInstance) {
        sortDialogInstance = new SortDialog(this.SN);
    }
    sortDialogInstance.show();
}
