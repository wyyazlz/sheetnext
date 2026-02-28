/**
 * CellIs规则 - 单元格值比较
 * 支持：大于、小于、等于、不等于、介于、不介于等
 */
import { CFRule } from '../CFRule.js';
import { compareValue } from '../helpers.js';

export class CellIsRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
        this.operator = config.operator || 'equal';
        this.formula1 = config.formula1;
        this.formula2 = config.formula2;
    }

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
