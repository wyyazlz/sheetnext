function _notify(SN, feature) {
    const msg = feature + ' features are disabled in SheetNext Open Edition.';
    if (SN?.Utils?.toast) SN.Utils.toast(msg);
    return false;
}

export function createPivotTable() { return _notify(this?.SN, 'PivotTable'); }
export function refreshPivotTable() { return _notify(this?.SN, 'PivotTable'); }
export function openPivotTableOptions() { return _notify(this?.SN, 'PivotTable'); }
export function showPivotReportFilterPages() { return _notify(this?.SN, 'PivotTable'); }
export function renamePivotTable() { return _notify(this?.SN, 'PivotTable'); }
export function pivotFieldSettings() { return _notify(this?.SN, 'PivotTable'); }
export function changePivotDataSource() { return _notify(this?.SN, 'PivotTable'); }
export function clearPivotTable() { return _notify(this?.SN, 'PivotTable'); }
export function movePivotTable() { return _notify(this?.SN, 'PivotTable'); }
export function insertPivotChart() { return _notify(this?.SN, 'PivotTable'); }
export function recommendPivotTable() { return _notify(this?.SN, 'PivotTable'); }
export function insertPivotSlicer() { return _notify(this?.SN, 'PivotTable'); }
export function togglePivotFieldList() { return _notify(this?.SN, 'PivotTable'); }
export function togglePivotDrillButtons() { return _notify(this?.SN, 'PivotTable'); }
export function togglePivotFieldHeaders() { return _notify(this?.SN, 'PivotTable'); }
export function expandPivotField() { return _notify(this?.SN, 'PivotTable'); }
export function collapsePivotField() { return _notify(this?.SN, 'PivotTable'); }
export function setPivotSubtotals() { return _notify(this?.SN, 'PivotTable'); }
export function setPivotGrandTotals() { return _notify(this?.SN, 'PivotTable'); }
export function setPivotLayout() { return _notify(this?.SN, 'PivotTable'); }
export function setPivotRepeatLabels() { return _notify(this?.SN, 'PivotTable'); }
export function setPivotBlankRows() { return _notify(this?.SN, 'PivotTable'); }
export function togglePivotOption() { return _notify(this?.SN, 'PivotTable'); }
