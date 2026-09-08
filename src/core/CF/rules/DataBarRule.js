/**
 * DataBar Rule - Data Bar
 * Support: gradient/solid color, positive and negative values, axis, minLength/maxLength
 */
import { CFRule } from '../CFRule.js';
import { resolveCfvoValue, expandHex } from '../helpers.js';

const _extensionIds = new WeakMap();

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
        /** @type {'automatic'|'middle'|'none'} */
        this.axisPosition = config.axisPosition || 'automatic';
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

        if (this.axisPosition === 'middle') {
            axisPosition = 0.5;
            percent = val < 0 ? (minVal < 0 ? val / minVal : 0) : (maxVal > 0 ? val / maxVal : 0);
        } else if (this.axisPosition === 'none') {
            axisPosition = null;
            percent = maxVal === minVal ? 0.5 : (val - minVal) / (maxVal - minVal);
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
            { '_$type': this.minCfvo.type === 'autoMin' ? 'min' : this.minCfvo.type },
            { '_$type': this.maxCfvo.type === 'autoMax' ? 'max' : this.maxCfvo.type }
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

        // Link the fallback rule to its Excel 2010 data bar extension.
        node.extLst = { ext: [{
            '_$uri': '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}',
            '_$xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
            'x14:id': this._getExtensionId()
        }] };
    }

    _getExtensionId() {
        if (!_extensionIds.has(this)) {
            const guid = globalThis.crypto?.randomUUID?.() || 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, ch => {
                const value = Math.floor(Math.random() * 16);
                return (ch === 'x' ? value : (value & 3) | 8).toString(16);
            });
            _extensionIds.set(this, `{${guid.toUpperCase()}}`);
        }
        return _extensionIds.get(this);
    }

    _toX14XmlObject() {
        const hasBorder = !!(this.borderColor || this.negativeBorderColor);
        const borderColor = this.borderColor || this.color;
        const negativeBorderColor = this.negativeBorderColor || borderColor;
        const sameNegativeColor = this.negativeColor.toUpperCase() === this.color.toUpperCase();
        const sameNegativeBorder = negativeBorderColor.toUpperCase() === borderColor.toUpperCase();
        const dataBar = {
            '_$minLength': String(this.minLength),
            '_$maxLength': String(this.maxLength),
            '_$showValue': this.showValue ? '1' : '0',
            '_$border': hasBorder ? '1' : '0',
            '_$gradient': this.gradient ? '1' : '0',
            '_$direction': this.direction,
            '_$negativeBarColorSameAsPositive': sameNegativeColor ? '1' : '0',
            '_$negativeBarBorderColorSameAsPositive': sameNegativeBorder ? '1' : '0',
            '_$axisPosition': this.axisPosition,
            'x14:cfvo': [this.minCfvo, this.maxCfvo].map(cfvo => {
                const value = { '_$type': cfvo.type };
                if (cfvo.val !== null && cfvo.val !== undefined && !['min', 'max', 'autoMin', 'autoMax'].includes(cfvo.type)) {
                    value['xm:f'] = String(cfvo.val).replace(/^=/, '');
                }
                return value;
            }),
            'x14:fillColor': { '_$rgb': expandHex(this.color) }
        };
        if (hasBorder) dataBar['x14:borderColor'] = { '_$rgb': expandHex(borderColor) };
        if (!sameNegativeColor) dataBar['x14:negativeFillColor'] = { '_$rgb': expandHex(this.negativeColor) };
        if (hasBorder && !sameNegativeBorder) dataBar['x14:negativeBorderColor'] = { '_$rgb': expandHex(negativeBorderColor) };
        if (this.axisPosition !== 'none') dataBar['x14:axisColor'] = { '_$rgb': expandHex(this.axisColor) };
        return {
            '_$type': 'dataBar',
            '_$id': this._getExtensionId(),
            'x14:dataBar': dataBar
        };
    }
}
