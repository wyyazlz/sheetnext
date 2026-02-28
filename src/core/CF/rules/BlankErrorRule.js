/**
 * BlankError规则 - 空白/非空白/错误/非错误
 */
import { CFRule } from '../CFRule.js';
import { isBlankValue, isErrorValue } from '../helpers.js';

export class BlankErrorRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
    }

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
