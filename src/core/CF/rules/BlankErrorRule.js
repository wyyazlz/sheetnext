/**
 * BlankError Rule - Blank/Not Blank/Error/Not Error
 */
import { CFRule } from '../CFRule.js';
import { isBlankValue, isErrorValue } from '../helpers.js';

export class BlankErrorRule extends CFRule {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        super(config, sheet);
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {boolean} */
    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;

        switch (this.type) {
            case 'containsBlanks':
                return isBlankValue(val);

            case 'notContainsBlanks':
                return !isBlankValue(val);

            case 'containsErrors':
                return isErrorValue(val);

            case 'notContainsErrors':
                return !isErrorValue(val);

            default:
                return false;
        }
    }

    _addXmlAttributes(node, cf) {
        // 这些类型没有额外属性
    }
}
