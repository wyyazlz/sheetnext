/**
 * Expression规则 - 自定义公式
 */
import { CFRule } from '../CFRule.js';

export class ExpressionRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
        this.formula = config.formula || '';
    }

    evaluate(cell, r, c, rangeData) {
        if (!this.formula) return false;

        // 计算公式结果（支持跨表引用）
        const result = this.resolveFormula(this.formula, r, c);

        // 公式结果为true或非零数值时满足条件
        if (typeof result === 'boolean') return result;
        if (typeof result === 'number') return result !== 0;

        return false;
    }

    _addXmlAttributes(node, cf) {
        if (this.formula) {
            node.formula = [String(this.formula)];
        }
    }
}
