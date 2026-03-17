/**
 * AboveAverage Rule - Above/Below Average
 */
import { CFRule } from '../CFRule.js';
import { getStdDev } from '../helpers.js';

export class AboveAverageRule extends CFRule {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        super(config, sheet);
        this.aboveAverage = config.aboveAverage !== false; // true=高于, false=低于
        /** @type {boolean} */
        this.equalAverage = config.equalAverage || false;  // 是否包含等于
        /** @type {number} */
        this.stdDev = config.stdDev || 0;                  // 标准差倍数
        this.needsRangeData = true;
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {boolean} */
    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (typeof val !== 'number' || isNaN(val)) return false;

        if (!rangeData) rangeData = this.getRangeData();

        const { avg, values } = rangeData;

        // 计算比较阈值
        let threshold = avg;
        if (this.stdDev > 0) {
            const numValues = values.map(v => v.val);
            const stdDev = getStdDev(numValues, avg);
            if (this.aboveAverage) {
                threshold = avg + stdDev * this.stdDev;
            } else {
                threshold = avg - stdDev * this.stdDev;
            }
        }

        // 比较
        if (this.aboveAverage) {
            if (this.equalAverage) {
                return val >= threshold;
            } else {
                return val > threshold;
            }
        } else {
            if (this.equalAverage) {
                return val <= threshold;
            } else {
                return val < threshold;
            }
        }
    }

    _addXmlAttributes(node, cf) {
        if (!this.aboveAverage) {
            node['_$aboveAverage'] = '0';
        }
        if (this.equalAverage) {
            node['_$equalAverage'] = '1';
        }
        if (this.stdDev > 0) {
            node['_$stdDev'] = String(this.stdDev);
        }
    }
}
