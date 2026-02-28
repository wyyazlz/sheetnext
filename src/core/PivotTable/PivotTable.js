function _notify(sheet) {
    const SN = sheet?.SN;
    const msg = 'PivotTable features are disabled in SheetNext Open Edition.';
    if (SN?.Utils?.toast) SN.Utils.toast(msg);
}

export default class PivotTable {
    constructor(sheet) {
        this.sheet = sheet;
        this.map = new Map();
        this._list = [];
    }

    get size() {
        return 0;
    }

    add() {
        _notify(this.sheet);
        return null;
    }

    remove() {
        return false;
    }

    get() {
        return null;
    }

    getAll() {
        return [];
    }

    refreshAll() {
        return false;
    }

    forEach() {}
    _addParsed() {}
    _remove() {}
}
