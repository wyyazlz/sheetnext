/**
 * Top10 Rule - Top N/Bottom N/Bottom N %/Bottom N %
 */
import { CFRule } from '../CFRule.js';

export class Top10Rule extends CFRule {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        super(config, sheet);
        /** @type {number} */
        this.rank = config.rank || 10;
        /** @type {boolean} */
        this.percent = config.percent || false;
        /** @type {boolean} */
        this.bottom = config.bottom || false;
        this.needsRangeData = true;
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {boolean} */
    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (typeof val !== 'number' || isNaN(val)) return false;

        if (!rangeData) rangeData = this.getRangeData();

        const { sorted, count } = rangeData;
        if (count === 0) return false;

        // 计算阈值数量
        let threshold;
        if (this.percent) {
            threshold = Math.max(1, Math.ceil(count * this.rank / 100));
        } else {
            threshold = Math.min(this.rank, count);
        }

        // 获取分界值
        let cutoffValue;
        if (this.bottom) {
            // 后N项：从最小开始
            cutoffValue = sorted[threshold - 1]?.val;
            return val <= cutoffValue;
        } else {
            // 前N项：从最大开始
            cutoffValue = sorted[count - threshold]?.val;
            return val >= cutoffValue;
        }
    }

    _addXmlAttributes(node, cf) {
        node['_$rank'] = String(this.rank);
        if (this.percent) {
            node['_$percent'] = '1';
        }
        if (this.bottom) {
            node['_$bottom'] = '1';
        }
    }
}
