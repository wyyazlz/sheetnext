const OPEN_PRINT_MESSAGE = 'Print module is not available in SheetNext Open Edition.';

function toastUnavailable(SN) {
    if (SN?.Utils?.toast) SN.Utils.toast(OPEN_PRINT_MESSAGE);
    return false;
}

export default class Print {
    constructor(SN) {
        this._SN = SN;
    }

    initSheetPrintSettings(sheet) {
        if (!sheet.printSettings) {
            sheet.printSettings = {
                showPageBreaks: false,
                showGridlines: false,
                showHeadings: false,
                pageMargins: {},
                pageSetup: {},
                printArea: null,
                printTitles: {},
                pageBackground: null,
                headerFooter: {}
            };
        }
        return sheet.printSettings;
    }

    ensureSettings(sheet) {
        return this.initSheetPrintSettings(sheet);
    }

    setSettings(partial) {
        const sheet = this._SN?.activeSheet;
        if (!sheet) return false;
        const settings = this.ensureSettings(sheet);
        Object.assign(settings, partial || {});
        return true;
    }

    setPageMargins(margins) {
        const sheet = this._SN?.activeSheet;
        if (!sheet) return false;
        const settings = this.ensureSettings(sheet);
        settings.pageMargins = { ...(settings.pageMargins || {}), ...(margins || {}) };
        return true;
    }

    setPageSetup(pageSetup) {
        const sheet = this._SN?.activeSheet;
        if (!sheet) return false;
        const settings = this.ensureSettings(sheet);
        settings.pageSetup = { ...(settings.pageSetup || {}), ...(pageSetup || {}) };
        return true;
    }

    setPrintArea(area) {
        const sheet = this._SN?.activeSheet;
        if (!sheet) return false;
        const settings = this.ensureSettings(sheet);
        settings.printArea = area || null;
        return true;
    }

    clearPrintArea() {
        return this.setPrintArea(null);
    }

    setPrintTitles(titles) {
        const sheet = this._SN?.activeSheet;
        if (!sheet) return false;
        const settings = this.ensureSettings(sheet);
        settings.printTitles = { ...(titles || {}) };
        return true;
    }

    preview() { return toastUnavailable(this._SN); }
    toggleShowPageBreaks() { return toastUnavailable(this._SN); }
    insertPageBreak() { return toastUnavailable(this._SN); }
    toggleGridlines() { return toastUnavailable(this._SN); }
    toggleHeadings() { return toastUnavailable(this._SN); }
    getPageBreakGuides() { return []; }
    applySettingsToWorksheet() { return false; }
}
