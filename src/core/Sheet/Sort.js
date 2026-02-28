import { sortRange } from '../Utils/SortUtils.js';

/** @param {Array<{col:string|number, order?:string, customOrder?:Array}>} sortKeys @param {RangeRef} [range] */
export function rangeSort(sortKeys, range) {
    const sheet = this;

    // 默认全表
    if (!range) {
        range = { s: { r: 0, c: 0 }, e: { r: sheet.rowCount - 1, c: sheet.colCount - 1 } };
    }

    const beforeEvent = sheet.SN.Event.emit('beforeSort', { sheet, range, sortKeys });
    if (beforeEvent.canceled) return false;

    return sheet.SN._withOperation('sort', { sheet, range, sortKeys }, () => {
        const success = sortRange(sheet, range, sortKeys);
        if (success) {
            sheet.SN._recordChange({ type: 'sort', sheet, range, sortKeys });
        }
        sheet.SN.Event.emit('afterSort', { sheet, range, sortKeys, success });
        return success;
    });
}
