function _notify(SN, feature) {
    const msg = feature + ' features are disabled in SheetNext Open Edition.';
    if (SN?.Utils?.toast) SN.Utils.toast(msg);
    return false;
}

export function printPreview() { return _notify(this?.SN, 'Print'); }
export function exportPdf() { return _notify(this?.SN, 'Print'); }
export function pageBreakPreview() { return _notify(this?.SN, 'Print'); }
export function toggleShowPageBreaks() { return _notify(this?.SN, 'Print'); }
export function insertPageBreak() { return _notify(this?.SN, 'Print'); }
export function togglePrintGridlines() { return _notify(this?.SN, 'Print'); }
export function togglePrintHeadings() { return _notify(this?.SN, 'Print'); }
export function setPageMarginsPreset() { return _notify(this?.SN, 'Print'); }
export function setPageMarginsDialog() { return _notify(this?.SN, 'Print'); }
export function setPageOrientation() { return _notify(this?.SN, 'Print'); }
export function setPaperSize() { return _notify(this?.SN, 'Print'); }
export function setPrintScale() { return _notify(this?.SN, 'Print'); }
export function setPrintScaleDialog() { return _notify(this?.SN, 'Print'); }
export function setPrintAreaFromSelection() { return _notify(this?.SN, 'Print'); }
export function clearPrintArea() { return _notify(this?.SN, 'Print'); }
export function setPrintTitles() { return _notify(this?.SN, 'Print'); }
export function setPageBackground() { return _notify(this?.SN, 'Print'); }
export function setHeaderFooter() { return _notify(this?.SN, 'Print'); }
