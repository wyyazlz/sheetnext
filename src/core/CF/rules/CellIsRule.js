/**
 * CellIs Rule - Cell Value Comparison
 * Support: greater than, less than, equal to, not equal to, between, not between etc.
 */
import { CFRule } from '../CFRule.js';
import { compareValue } from '../helpers.js';

export class CellIsRule extends CFRule {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        super(config, sheet);
        /** @type {string} */
        this.operator = config.operator || 'equal';
        /** @type {string|number|null|undefined} */
        this.formula1 = config.formula1;
        /** @type {string|number|null|undefined} */
        this.formula2 = config.formula2;
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {boolean} */
    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;

        // 解析公式值（支持跨表引用）
        const f1 = this.resolveFormula(this.formula1, r, c);
        const f2 = this.resolveFormula(this.formula2, r, c);

        // 数值比较
        if (typeof val === 'number' && typeof f1 === 'number') {
            return compareValue(val, this.operator, f1, f2);
        }

        // 字符串比较（用于equal/notEqual）
        if (this.operator === 'equal') {
            return String(val) === String(f1);
        }
        if (this.operator === 'notEqual') {
            return String(val) !== String(f1);
        }

        return false;
    }

    _addXmlAttributes(node, cf) {
        node['_$operator'] = this.operator;

        const formulas = [];
        if (this.formula1 !== null && this.formula1 !== undefined) {
            formulas.push(String(this.formula1));
        }
        if (this.formula2 !== null && this.formula2 !== undefined) {
            formulas.push(String(this.formula2));
        }
        if (formulas.length > 0) {
            node.formula = formulas;
        }
    }
}
