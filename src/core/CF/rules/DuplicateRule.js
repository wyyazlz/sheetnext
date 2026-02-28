/**
 * Duplicate规则 - 重复值/唯一值
 */
import { CFRule } from '../CFRule.js';

export class DuplicateRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
        this.needsRangeData = true;
    }

    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (val === null || val === undefined || val === '') return false;

        if (!rangeData) rangeData = this.getRangeData();

        const key = String(val);
        const occurrences = rangeData.cellMap.get(key);

        if (!occurrences) return false;

        const isDuplicate = occurrences.length > 1;

        if (this.type === 'duplicateValues') {
            return isDuplicate;
        } else if (this.type === 'uniqueValues') {
            return !isDuplicate;
        }

        return false;
    }

    _addXmlAttributes(node, cf) {
        // duplicateValues 和 uniqueValues 没有额外属性
    }
}
