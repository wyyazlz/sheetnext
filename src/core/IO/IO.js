import * as JsonIO from './JsonIO.js';

const _lm = new WeakMap();
export const _gl = (io) => _lm.get(io);

const OPEN_IO_MESSAGE = 'IO import/export is not available in SheetNext Open Edition.';

function notifyUnavailable(SN) {
    if (SN?.Utils?.toast) SN.Utils.toast(OPEN_IO_MESSAGE);
    return false;
}

export default class IO {
    /**
     * @param {import('../Workbook/Workbook.js').default} SN
     * @param {import('../License/License.js').default} [license]
     */
    constructor(SN, license) {
        /** @type {import('../Workbook/Workbook.js').default} */
        this._SN = SN;
        _lm.set(this, license);
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
