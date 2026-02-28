/**
 * @title 🎯 条件格式
 */
import { createRule } from './ruleFactory.js';

/**
 * CF 条件格式管理器
 * 负责解析、存储和管理 Excel 条件格式规则
 */
export default class CF {
    /**
     * @param {Sheet} sheet - 所属工作表
     */
    constructor(sheet) {
        /** @type {Sheet} */
        this.sheet = sheet;
        /** @type {Array<CFRule>} 规则列表 */
        this.rules = [];
        // 缓存
        this._formatCache = new Map();
        this._rangeDataCache = new Map();
        // dxf样式管理（用于导出）
        this._dxfList = [];
        this._dxfMap = new Map();
    }

    // ==================== 增删改查 ====================

    /**
     * 添加条件格式规则
     * @param {Object} config - 规则配置
     * @param {string} config.rangeRef - 应用范围 "A1:D10"
     * @param {string} config.type - 规则类型 (colorScale/dataBar/iconSet/cellIs/expression/top10/aboveAverage/duplicateValues/containsText/timePeriod/containsBlanks/containsErrors)
     * @returns {CFRule} 创建的规则
     */
    add(config = {}) {
        const { rangeRef } = config;
        if (!rangeRef) return null;

        const ruleConfig = {
            ...config,
            sqref: rangeRef,
            priority: config.priority ?? this._getNextPriority()
        };
        const rule = createRule(ruleConfig, this.sheet);
        this.rules.push(rule);
        this._sortRules();
        this.clearRangeCache(rangeRef);
        return rule;
    }

    /**
     * 删除条件格式规则
     * @param {number} index - 规则索引
     * @returns {boolean}
     */
    remove(index) {
        if (index < 0 || index >= this.rules.length) return false;
        const rule = this.rules[index];
        this.rules.splice(index, 1);
        this.clearRangeCache(rule.sqref);
        return true;
    }

    /**
     * 获取规则
     * @param {number} index - 规则索引
     * @returns {CFRule|null}
     */
    get(index) {
        return this.rules[index] ?? null;
    }

    /**
     * 获取所有规则
     * @returns {Array<CFRule>}
     */
    getAll() {
        return this.rules;
    }

    /**
     * 获取下一个优先级
     * @private
     */
    _getNextPriority() {
        if (this.rules.length === 0) return 1;
        return Math.max(...this.rules.map(r => r.priority)) + 1;
    }

    // ==================== 解析 ====================

    /**
     * 从 xmlObj 解析条件格式数据
     * @param {Object} xmlObj - Sheet 的 XML 对象
     * @returns {void}
     */
    parse(xmlObj) {
        const cfList = xmlObj?.conditionalFormatting;
        if (!cfList) return;

        const items = Array.isArray(cfList) ? cfList : [cfList];
        const dxfs = this.sheet.SN.Xml.obj['xl/styles.xml']?.styleSheet?.dxfs?.dxf || [];
        const dxfArr = Array.isArray(dxfs) ? dxfs : [dxfs];

        for (const cf of items) {
            const sqref = cf['_$sqref'];
            const cfRules = cf.cfRule;
            if (!cfRules) continue;

            const ruleList = Array.isArray(cfRules) ? cfRules : [cfRules];
            for (const ruleXml of ruleList) {
                const config = this._parseRuleXml(ruleXml, sqref, dxfArr);
                const rule = createRule(config, this.sheet);
                this.rules.push(rule);
            }
        }

        this._sortRules();
    }

    /**
     * 解析单个规则XML
     * @private
     */
    _parseRuleXml(ruleXml, sqref, dxfArr) {
        const config = {
            type: ruleXml['_$type'],
            priority: parseInt(ruleXml['_$priority']) || 1,
            sqref,
            stopIfTrue: ruleXml['_$stopIfTrue'] === '1'
        };

        // 解析dxf样式
        const dxfId = ruleXml['_$dxfId'];
        if (dxfId !== undefined) {
            config.dxf = this._parseDxf(dxfArr[parseInt(dxfId)]);
        }

        // 解析公式
        if (ruleXml.formula) {
            const formulas = Array.isArray(ruleXml.formula) ? ruleXml.formula : [ruleXml.formula];
            config.formula1 = formulas[0];
            config.formula2 = formulas[1];
        }

        // 解析操作符
        if (ruleXml['_$operator']) {
            config.operator = ruleXml['_$operator'];
        }

        // 解析文本
        if (ruleXml['_$text']) {
            config.text = ruleXml['_$text'];
        }

        // 解析timePeriod
        if (ruleXml['_$timePeriod']) {
            config.timePeriod = ruleXml['_$timePeriod'];
        }

        // 解析aboveAverage
        if (ruleXml['_$aboveAverage'] !== undefined) {
            config.aboveAverage = ruleXml['_$aboveAverage'] !== '0';
        }
        if (ruleXml['_$equalAverage'] !== undefined) {
            config.equalAverage = ruleXml['_$equalAverage'] === '1';
        }
        if (ruleXml['_$stdDev'] !== undefined) {
            config.stdDev = parseInt(ruleXml['_$stdDev']);
        }

        // 解析top10
        if (ruleXml['_$rank'] !== undefined) {
            config.rank = parseInt(ruleXml['_$rank']);
        }
        if (ruleXml['_$percent'] !== undefined) {
            config.percent = ruleXml['_$percent'] === '1';
        }
        if (ruleXml['_$bottom'] !== undefined) {
            config.bottom = ruleXml['_$bottom'] === '1';
        }

        // 解析colorScale
        if (ruleXml.colorScale) {
            const colorScale = this._parseColorScale(ruleXml.colorScale);
            if (colorScale) {
                config.cfvos = colorScale.cfvo;
                config.colors = colorScale.colors;
            }
        }

        // 解析dataBar
        if (ruleXml.dataBar) {
            const dataBar = this._parseDataBar(ruleXml.dataBar);
            if (dataBar) {
                const cfvoArr = Array.isArray(dataBar.cfvo) ? dataBar.cfvo : [];
                config.minCfvo = cfvoArr[0] || { type: 'min' };
                config.maxCfvo = cfvoArr[1] || { type: 'max' };
                config.color = dataBar.color;
                config.showValue = dataBar.showValue;
                config.minLength = dataBar.minLength;
                config.maxLength = dataBar.maxLength;
            }
        }

        // 解析iconSet
        if (ruleXml.iconSet) {
            const iconSet = this._parseIconSet(ruleXml.iconSet);
            if (iconSet) {
                config.iconSet = iconSet.iconSet;
                config.cfvos = iconSet.cfvo;
                config.showValue = iconSet.showValue;
                config.reverse = iconSet.reverse;
            }
        }

        return config;
    }

    /**
     * 解析dxf样式
     * @private
     */
    _parseDxf(dxfXml) {
        if (!dxfXml) return null;

        const dxf = {};

        // 字体
        if (dxfXml.font) {
            dxf.font = {};
            if (dxfXml.font.b) dxf.font.bold = true;
            if (dxfXml.font.i) dxf.font.italic = true;
            if (dxfXml.font.strike) dxf.font.strike = true;
            if (dxfXml.font.u) dxf.font.underline = true;
            if (dxfXml.font.color) {
                dxf.font.color = this._parseColor(dxfXml.font.color);
            }
        }

        // 填充
        if (dxfXml.fill?.patternFill) {
            dxf.fill = {};
            const pf = dxfXml.fill.patternFill;
            if (pf.bgColor) {
                dxf.fill.fgColor = this._parseColor(pf.bgColor);
            } else if (pf.fgColor) {
                dxf.fill.fgColor = this._parseColor(pf.fgColor);
            }
        }

        // 边框
        if (dxfXml.border) {
            dxf.border = {};
            const sides = ['top', 'bottom', 'left', 'right'];
            for (const side of sides) {
                if (dxfXml.border[side]) {
                    dxf.border[side] = {
                        style: dxfXml.border[side]['_$style'] || 'thin',
                        color: dxfXml.border[side].color ? this._parseColor(dxfXml.border[side].color) : '#000000'
                    };
                }
            }
        }

        return Object.keys(dxf).length > 0 ? dxf : null;
    }

    /**
     * 解析颜色
     * @private
     */
    _parseColor(colorXml) {
        if (!colorXml) return null;
        if (colorXml['_$rgb']) {
            const rgb = colorXml['_$rgb'];
            return '#' + (rgb.length === 8 ? rgb.slice(2) : rgb);
        }
        if (colorXml['_$theme'] !== undefined) {
            return this.sheet.SN.Xml.getThemeColor(parseInt(colorXml['_$theme']), parseFloat(colorXml['_$tint'] || 0));
        }
        if (colorXml['_$indexed'] !== undefined) {
            return this.sheet.SN.Xml.getIndexedColor(parseInt(colorXml['_$indexed']));
        }
        return null;
    }

    /**
     * 解析colorScale
     * @private
     */
    _parseColorScale(csXml) {
        const cfvo = csXml.cfvo;
        const color = csXml.color;
        if (!cfvo || !color) return null;

        const cfvoArr = Array.isArray(cfvo) ? cfvo : [cfvo];
        const colorArr = Array.isArray(color) ? color : [color];

        return {
            cfvo: cfvoArr.map(c => ({
                type: c['_$type'],
                val: c['_$val'] !== undefined ? parseFloat(c['_$val']) : undefined
            })),
            colors: colorArr.map(c => this._parseColor(c))
        };
    }

    /**
     * 解析dataBar
     * @private
     */
    _parseDataBar(dbXml) {
        const cfvo = dbXml.cfvo;
        if (!cfvo) return null;

        const cfvoArr = Array.isArray(cfvo) ? cfvo : [cfvo];

        return {
            cfvo: cfvoArr.map(c => ({
                type: c['_$type'],
                val: c['_$val'] !== undefined ? parseFloat(c['_$val']) : undefined
            })),
            color: dbXml.color ? this._parseColor(dbXml.color) : '#638EC6',
            showValue: dbXml['_$showValue'] !== '0',
            minLength: parseInt(dbXml['_$minLength']) || 10,
            maxLength: parseInt(dbXml['_$maxLength']) || 90
        };
    }

    /**
     * 解析iconSet
     * @private
     */
    _parseIconSet(isXml) {
        const cfvo = isXml.cfvo;
        if (!cfvo) return null;

        const cfvoArr = Array.isArray(cfvo) ? cfvo : [cfvo];

        return {
            iconSet: isXml['_$iconSet'] || '3TrafficLights1',
            cfvo: cfvoArr.map(c => ({
                type: c['_$type'],
                val: c['_$val'] !== undefined ? parseFloat(c['_$val']) : undefined,
                gte: c['_$gte'] !== '0'
            })),
            showValue: isXml['_$showValue'] !== '0',
            reverse: isXml['_$reverse'] === '1'
        };
    }

    // ==================== 公共方法 ====================

    /**
     * 获取单元格的条件格式
     * @param {number} r - 行索引
     * @param {number} c - 列索引
     * @returns {Object|null} 格式对象
     */
    getFormat(r, c) {
        const key = `${r}:${c}`;
        if (this._formatCache.has(key)) {
            return this._formatCache.get(key);
        }

        const cell = this.sheet.getCell(r, c);
        if (!cell) return null;

        let result = null;

        for (const rule of this.rules) {
            if (!this._isCellInRange(r, c, rule.sqref)) continue;

            const rangeData = rule.needsRangeData ? this._getRangeData(rule.sqref) : null;
            const format = rule.getFormat(cell, r, c, rangeData);

            if (format) {
                result = result ? this._mergeFormat(result, format) : format;
                if (rule.stopIfTrue) break;
            }
        }

        this._formatCache.set(key, result);
        return result;
    }

    /**
     * 清除所有缓存
     * @returns {void}
     */
    clearAllCache() {
        this._formatCache.clear();
        this._rangeDataCache.clear();
        for (const rule of this.rules) {
            rule._clearCache();
        }
    }

    /**
     * 清除指定范围的缓存
     * @param {string} sqref - 范围
     * @returns {void}
     */
    clearRangeCache(sqref) {
        this._rangeDataCache.delete(sqref);

        const ranges = sqref.split(' ');
        for (const rangeStr of ranges) {
            if (!rangeStr) continue;
            const range = this.sheet.rangeStrToNum(rangeStr);
            for (const key of this._formatCache.keys()) {
                const [r, c] = key.split(':').map(Number);
                if (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c) {
                    this._formatCache.delete(key);
                }
            }
        }
    }

    // ==================== 导出 ====================

    /**
     * 导出为worksheet XML节点数组
     * @returns {Array}
     */
    toXmlObject() {
        if (this.rules.length === 0) return null;

        // 按sqref分组
        const groups = new Map();
        for (const rule of this.rules) {
            if (!groups.has(rule.sqref)) {
                groups.set(rule.sqref, []);
            }
            groups.get(rule.sqref).push(rule);
        }

        const result = [];
        groups.forEach((rules, sqref) => {
            rules.sort((a, b) => a.priority - b.priority);
            result.push({
                '_$sqref': sqref,
                cfRule: rules.map(r => r._toXmlObject(this))
            });
        });

        return result.length > 0 ? result : null;
    }

    /**
     * 获取dxf样式索引（用于导出）
     * @param {Object} dxf - 差异化样式对象
     * @returns {number} dxfId
     */
    getDxfId(dxf) {
        const key = JSON.stringify(dxf);
        if (this._dxfMap.has(key)) {
            return this._dxfMap.get(key);
        }
        const id = this._dxfList.length;
        this._dxfList.push(dxf);
        this._dxfMap.set(key, id);
        return id;
    }

    /**
     * 获取dxf样式列表（用于导出styles.xml）
     * @returns {Array}
     */
    getDxfList() {
        return this._dxfList;
    }

    // ==================== 私有方法 ====================

    _isCellInRange(r, c, sqref) {
        const ranges = sqref.split(' ');
        for (const rangeStr of ranges) {
            const range = this.sheet.rangeStrToNum(rangeStr);
            if (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c) {
                return true;
            }
        }
        return false;
    }

    _getRangeData(sqref) {
        if (this._rangeDataCache.has(sqref)) {
            return this._rangeDataCache.get(sqref);
        }
        // 使用规则的getRangeData方法
        const rule = this.rules.find(r => r.sqref === sqref);
        if (rule) {
            const data = rule.getRangeData();
            this._rangeDataCache.set(sqref, data);
            return data;
        }
        return null;
    }

    _sortRules() {
        this.rules.sort((a, b) => a.priority - b.priority);
    }

    _mergeFormat(base, overlay) {
        const result = { ...base };

        if (overlay.fill) result.fill = overlay.fill;
        if (overlay.font) result.font = { ...base.font, ...overlay.font };
        if (overlay.border) result.border = { ...base.border, ...overlay.border };
        if (overlay.dataBar && !base.dataBar) result.dataBar = overlay.dataBar;
        if (overlay.icon && !base.icon) result.icon = overlay.icon;
        if (overlay.colorScale && !base.colorScale) result.colorScale = overlay.colorScale;

        return result;
    }
}
