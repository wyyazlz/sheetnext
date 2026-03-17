/**
 * Duplicate Rule - Duplicate Value/Unique Value
 */
import { CFRule } from '../CFRule.js';

export class DuplicateRule extends CFRule {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        super(config, sheet);
        this.needsRangeData = true;
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {boolean} */
    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (val === null || val === undefined || val === '') return false;

        if (!rangeData) rangeData = this.getRangeData();

        const key = String(val);
        const occurrences = rangeData.cellMap.get(key);

        if (!occurrences) return false;

        const isDuplicate = occurrences.length > 1;
        const type = this.type ?? this._config.type;

        if (type === 'duplicateValues') {
            return isDuplicate;
        } else if (type === 'uniqueValues') {
            return !isDuplicate;
        }

        return false;
    }

    _addXmlAttributes(node, cf) {
        // duplicateValues 和 uniqueValues 没有额外属性
    }
}
