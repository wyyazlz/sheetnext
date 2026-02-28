const OPEN_DRAWING_MESSAGE = 'Drawing module is not available in SheetNext Open Edition.';

function notifyDisabled(sheet) {
    const SN = sheet?.SN;
    if (SN?.Utils?.toast) SN.Utils.toast(OPEN_DRAWING_MESSAGE);
}

export default class Drawing {
    constructor(sheet) {
        this.sheet = sheet;
    }

    addChart() {
        notifyDisabled(this.sheet);
        return null;
    }

    addImage() {
        notifyDisabled(this.sheet);
        return null;
    }

    addShape() {
        notifyDisabled(this.sheet);
        return null;
    }

    _add() {
        return null;
    }

    remove() {
        return false;
    }

    get() {
        return null;
    }

    getById() {
        return null;
    }

    getAtCell() {
        return [];
    }

    getAll() {
        return [];
    }

    _parse() {}
    _render() {}
    _onRowsInserted() {}
    _onColsInserted() {}
    _onRowsDeleted() {}
    _onColsDeleted() {}
    _clearChartRefCaches() {}
}
