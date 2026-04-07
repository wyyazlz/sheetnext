import { sortRange } from '../Utils/SortUtils.js';

/** @param {Array<{col:string|number, order?:string, customOrder?:Array}>} sortKeys @param {RangeRef} [range] @param {{hasHeader?:boolean, caseSensitive?:boolean}} [options] @returns {boolean} Sort sheet range. */
export function rangeSort(sortKeys, range, options = {}) {
    const sheet = this;

    if (!range) {
        range = { s: { r: 0, c: 0 }, e: { r: sheet.rowCount - 1, c: sheet.colCount - 1 } };
    }

    const beforeEvent = sheet.SN.Event.emit('beforeSort', { sheet, range, sortKeys, options });
    if (beforeEvent.canceled) return false;

    return sheet.SN._withOperation('sort', { sheet, range, sortKeys, options }, () => {
        const success = sortRange(sheet, range, sortKeys, options);
        if (success) {
            sheet.SN._recordChange({ type: 'sort', sheet, range, sortKeys, options });
        }
        sheet.SN.Event.emit('afterSort', { sheet, range, sortKeys, options, success });
        return success;
    });
}
