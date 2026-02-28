import * as JsonIO from './JsonIO.js';

const OPEN_IO_MESSAGE = 'IO import/export is not available in SheetNext Open Edition.';

function notifyUnavailable(SN) {
    if (SN?.Utils?.toast) SN.Utils.toast(OPEN_IO_MESSAGE);
    return false;
}

export default class IO {
    /** @param {import('../Workbook/Workbook.js').default} SN */
    constructor(SN) {
        /** @type {import('../Workbook/Workbook.js').default} */
        this._SN = SN;
    }

    async import() {
        return notifyUnavailable(this._SN);
    }

    async importFromUrl() {
        return notifyUnavailable(this._SN);
    }

    export() {
        return notifyUnavailable(this._SN);
    }

    exportAllImage() {
        return notifyUnavailable(this._SN);
    }
}

Object.assign(IO.prototype, {
    getData: JsonIO.getData,
    setData: JsonIO.setData,
    _setData: JsonIO.setData
});
