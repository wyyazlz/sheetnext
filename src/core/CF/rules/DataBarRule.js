/**
 * DataBar Rule - Data Bar
 * Support: gradient/solid color, positive and negative values, axis, minLength/maxLength
 */
import { CFRule } from '../CFRule.js';
import { resolveCfvoValue, expandHex } from '../helpers.js';

export class DataBarRule extends CFRule {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        super(config, sheet);
        /** @type {Object} */
        this.minCfvo = config.minCfvo || { type: 'min' };
        /** @type {Object} */
        this.maxCfvo = config.maxCfvo || { type: 'max' };
        /** @type {string} */
        this.color = config.color || '#638EC6';
        this.showValue = config.showValue !== false;
        this.gradient = config.gradient !== false;
        /** @type {string} */
        this.negativeColor = config.negativeColor || '#FF0000';
        /** @type {string|null} */
        this.borderColor = config.borderColor || null;
        /** @type {string|null} */
        this.negativeBorderColor = config.negativeBorderColor || null;
        /** @type {string} */
        this.axisColor = config.axisColor || '#000000';
        /** @type {string} */
        this.direction = config.direction || 'context';
        // Excel默认: minLength=10%, maxLength=90%
        /** @type {number} */
        this.minLength = config.minLength ?? 10;
        /** @type {number} */
        this.maxLength = config.maxLength ?? 90;
        this.needsRangeData = true;
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {boolean} */
    evaluate(cell, r, c, rangeData) {
        return typeof cell.calcVal === 'number' && !isNaN(cell.calcVal);
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {Object|null} */
    getFormat(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (typeof val !== 'number' || isNaN(val)) return null;

        if (!rangeData) rangeData = this.getRangeData();

        const minVal = resolveCfvoValue(this.minCfvo, rangeData, (f) => this.resolveFormula(f, r, c));
        const maxVal = resolveCfvoValue(this.maxCfvo, rangeData, (f) => this.resolveFormula(f, r, c));

        const hasNegative = minVal < 0;
        const hasPositive = maxVal > 0;

        let percent;
        let axisPosition = null;

        if (hasNegative && hasPositive) {
            const range = maxVal - minVal;
            axisPosition = Math.abs(minVal) / range;

            if (val >= 0) {
                percent = val / maxVal;
            } else {
                percent = val / minVal;
            }
        } else {
            if (maxVal === minVal) {
                percent = 0.5;
            } else {
                percent = (val - minVal) / (maxVal - minVal);
            }
        }

        percent = Math.max(-1, Math.min(1, percent));

        // 应用 minLength/maxLength 缩放
        const minL = this.minLength / 100;
        const maxL = this.maxLength / 100;
        const scaledPercent = minL + (maxL - minL) * Math.abs(percent);

        return {
            dataBar: {
                percent: val < 0 ? -scaledPercent : scaledPercent,
                color: val >= 0 ? this.color : this.negativeColor,
                gradient: this.gradient,
                borderColor: val >= 0 ? this.borderColor : this.negativeBorderColor,
                showValue: this.showValue,
                axisPosition,
                axisColor: this.axisColor,
                direction: this.direction,
                isNegative: val < 0
            }
        };
    }

    _addXmlAttributes(node, cf) {
        const cfvos = [
            { '_$type': this.minCfvo.type },
            { '_$type': this.maxCfvo.type }
        ];
        if (this.minCfvo.val !== null && this.minCfvo.val !== undefined) {
            cfvos[0]['_$val'] = String(this.minCfvo.val);
        }
        if (this.maxCfvo.val !== null && this.maxCfvo.val !== undefined) {
            cfvos[1]['_$val'] = String(this.maxCfvo.val);
        }

        node.dataBar = {
            cfvo: cfvos,
            color: { '_$rgb': expandHex(this.color) }
        };

        if (!this.showValue) {
            node.dataBar['_$showValue'] = '0';
        }
        if (this.minLength !== 10) {
            node.dataBar['_$minLength'] = String(this.minLength);
        }
        if (this.maxLength !== 90) {
            node.dataBar['_$maxLength'] = String(this.maxLength);
        }
    }
}
