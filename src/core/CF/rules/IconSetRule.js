/**
 * IconSet规则 - 图标集
 * 支持20+种图标集
 */
import { CFRule } from '../CFRule.js';
import { resolveCfvoValue } from '../helpers.js';

// 图标集配置（index 0 = 最低值图标，index n-1 = 最高值图标）
export const ICON_SETS = {
    '3Arrows': { count: 3, icons: ['arrowDown', 'arrowSide', 'arrowUp'] },
    '3ArrowsGray': { count: 3, icons: ['arrowDownGray', 'arrowSideGray', 'arrowUpGray'] },
    '3Flags': { count: 3, icons: ['flagRed', 'flagYellow', 'flagGreen'] },
    '3Signs': { count: 3, icons: ['signRed', 'signYellow', 'signGreen'] },
    '3Symbols': { count: 3, icons: ['symbolX', 'symbolExclaim', 'symbolCheck'] },
    '3Symbols2': { count: 3, icons: ['symbolX2', 'symbolExclaim2', 'symbolCheck2'] },
    '3TrafficLights1': { count: 3, icons: ['lightRed', 'lightYellow', 'lightGreen'] },
    '3TrafficLights2': { count: 3, icons: ['lightRed2', 'lightYellow2', 'lightGreen2'] },
    '3Stars': { count: 3, icons: ['starEmpty', 'starHalf', 'starFull'] },
    '3Triangles': { count: 3, icons: ['triangleDown', 'triangleDash', 'triangleUp'] },
    '4Arrows': { count: 4, icons: ['arrowDown', 'arrowTiltDown', 'arrowTiltUp', 'arrowUp'] },
    '4ArrowsGray': { count: 4, icons: ['arrowDownGray', 'arrowTiltDownGray', 'arrowTiltUpGray', 'arrowUpGray'] },
    '4Rating': { count: 4, icons: ['rating1', 'rating2', 'rating3', 'rating4'] },
    '4RedToBlack': { count: 4, icons: ['circleBlack', 'circleGray', 'circlePink', 'circleRed'] },
    '4TrafficLights': { count: 4, icons: ['lightBlack', 'lightRed', 'lightYellow', 'lightGreen'] },
    '5Arrows': { count: 5, icons: ['arrowDown', 'arrowTiltDown', 'arrowSide', 'arrowTiltUp', 'arrowUp'] },
    '5ArrowsGray': { count: 5, icons: ['arrowDownGray', 'arrowTiltDownGray', 'arrowSideGray', 'arrowTiltUpGray', 'arrowUpGray'] },
    '5Boxes': { count: 5, icons: ['box0', 'box1', 'box2', 'box3', 'box4'] },
    '5Quarters': { count: 5, icons: ['quarter0', 'quarter1', 'quarter2', 'quarter3', 'quarter4'] },
    '5Rating': { count: 5, icons: ['rating1', 'rating2', 'rating3', 'rating4', 'rating5'] }
};

export class IconSetRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
        const parsedIconSet = (config.iconSet && typeof config.iconSet === 'object') ? config.iconSet : null;
        this.iconSet = (typeof config.iconSet === 'string' ? config.iconSet : parsedIconSet?.iconSet) || '3TrafficLights1';
        this.cfvos = config.cfvos || parsedIconSet?.cfvo || this._getDefaultCfvos();
        this.reverse = (typeof config.reverse === 'boolean' ? config.reverse : parsedIconSet?.reverse) || false;
        const showValue = typeof config.showValue !== 'undefined' ? config.showValue : parsedIconSet?.showValue;
        this.showValue = showValue !== false;
        this.needsRangeData = true;
    }

    _getDefaultCfvos() {
        const setConfig = ICON_SETS[this.iconSet];
        if (!setConfig) return [{ type: 'percent', val: 0 }];

        const count = setConfig.count;
        const cfvos = [];
        for (let i = 0; i < count; i++) {
            cfvos.push({ type: 'percent', val: Math.round(i * 100 / count) });
        }
        return cfvos;
    }

    evaluate(cell, r, c, rangeData) {
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

        // 确定图标索引
        let iconIndex = this._getIconIndex(val, thresholds);

        // 反转
        if (this.reverse) {
            const setConfig = ICON_SETS[this.iconSet];
            if (setConfig) {
                iconIndex = setConfig.count - 1 - iconIndex;
            }
        }

        return {
            icon: {
                set: this.iconSet,
                index: iconIndex,
                showValue: this.showValue
            }
        };
    }

    _getIconIndex(val, thresholds) {
        // 从高到低匹配
        for (let i = thresholds.length - 1; i >= 0; i--) {
            if (val >= thresholds[i]) {
                return i;
            }
        }
        return 0;
    }

    _addXmlAttributes(node, cf) {
        const iconSetName = typeof this.iconSet === 'string' ? this.iconSet : '3TrafficLights1';
        node.iconSet = {
            '_$iconSet': iconSetName,
            cfvo: this.cfvos.map(cfvo => {
                const obj = { '_$type': cfvo.type };
                if (cfvo.val !== null && cfvo.val !== undefined) {
                    obj['_$val'] = String(cfvo.val);
                }
                return obj;
            })
        };

        if (this.reverse) {
            node.iconSet['_$reverse'] = '1';
        }
        if (!this.showValue) {
            node.iconSet['_$showValue'] = '0';
        }
    }
}
