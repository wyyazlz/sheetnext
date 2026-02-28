/**
 * ColorScale规则 - 色阶（双色/三色渐变）
 */
import { CFRule } from '../CFRule.js';
import { interpolateColor, interpolateColor3, resolveCfvoValue, expandHex } from '../helpers.js';

export class ColorScaleRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
        this.cfvos = config.cfvos || [{ type: 'min' }, { type: 'max' }];
        this.colors = config.colors || ['#F8696B', '#63BE7B'];
        this.needsRangeData = true;
    }

    evaluate(cell, r, c, rangeData) {
        // 色阶只对数值有效
        return typeof cell.calcVal === 'number' && !isNaN(cell.calcVal);
    }

    getFormat(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (typeof val !== 'number' || isNaN(val)) return null;

        if (!rangeData) rangeData = this.getRangeData();

        // 计算阈值
        const thresholds = this.cfvos.map(cfvo =>
            resolveCfvoValue(cfvo, rangeData, (f) => this.resolveFormula(f, r, c))
        );

        // 计算颜色
        const color = this._calculateColor(val, thresholds);

        return {
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: color
            },
            _colorScale: true // 标记为色阶（渲染时可能有特殊处理）
        };
    }

    _calculateColor(val, thresholds) {
        const colors = this.colors;

        if (colors.length === 2) {
            // 双色渐变
            const [min, max] = thresholds;
            if (max === min) return colors[0];
            const t = Math.max(0, Math.min(1, (val - min) / (max - min)));
            return interpolateColor(colors[0], colors[1], t);
        } else if (colors.length >= 3) {
            // 三色渐变
            const [min, mid, max] = thresholds;
            let t;
            if (val <= mid) {
                t = min === mid ? 0 : (val - min) / (mid - min) * 0.5;
            } else {
                t = mid === max ? 1 : 0.5 + (val - mid) / (max - mid) * 0.5;
            }
            t = Math.max(0, Math.min(1, t));
            return interpolateColor3(colors[0], colors[1], colors[2], t);
        }

        return colors[0] || '#FFFFFF';
    }

    _addXmlAttributes(node, cf) {
        node.colorScale = {
            cfvo: this.cfvos.map(cfvo => {
                const obj = { '_$type': cfvo.type };
                if (cfvo.val !== null && cfvo.val !== undefined) {
                    obj['_$val'] = String(cfvo.val);
                }
                return obj;
            }),
            color: this.colors.map(c => ({ '_$rgb': expandHex(c) }))
        };
    }
}
