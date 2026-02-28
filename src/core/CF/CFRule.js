/** @class */
export class CFRule {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        this.sheet = sheet;
        /** @type {string} */
        this.type = config.type;
        /** @type {number} */
        this.priority = config.priority;
        /** @type {string} */
        this.sqref = config.sqref;
        /** @type {boolean} */
        this.stopIfTrue = config.stopIfTrue || false;
        /** @type {Object|null} */
        this.dxf = config.dxf || null;
        /** @type {boolean} */
        this.needsRangeData = false;
        this._SN = sheet.SN;
        this._config = { ...config };
        this._evalCache = new Map();
        this._rangeDataCache = null;
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {boolean} */
    evaluate(cell, r, c, rangeData) {
        return false;
    }

    /** @param {Cell} cell @param {number} r @param {number} c @param {Object} rangeData @returns {Object|null} */
    getFormat(cell, r, c, rangeData) {
        // 对于使用dxf的规则，先评估条件
        if (!this.evaluate(cell, r, c, rangeData)) {
            return null;
        }
        // 返回dxf样式
        return this.dxf ? { ...this.dxf } : null;
    }

    /** @returns {Object} */
    getRangeData() {
        if (this._rangeDataCache) return this._rangeDataCache;

        const range = this.sheet.rangeStrToNum(this.sqref);
        const values = [];
        const textValues = [];
        const cellMap = new Map();

        // 稀疏遍历：只访问已存在的行和单元格
        const rows = this.sheet.rows;
        for (const rIndex in rows) {
            const r = parseInt(rIndex);
            if (r < range.s.r || r > range.e.r) continue;
            const row = rows[rIndex];
            if (!row.cells) continue;
            for (const cIndex in row.cells) {
                const c = parseInt(cIndex);
                if (c < range.s.c || c > range.e.c) continue;
                const cell = row.cells[cIndex];
                const val = cell.calcVal;

                if (typeof val === 'number' && !isNaN(val)) {
                    values.push({ r, c, val });
                }

                const strVal = String(val ?? '');
                if (!cellMap.has(strVal)) {
                    cellMap.set(strVal, []);
                }
                cellMap.get(strVal).push({ r, c });

                if (val !== null && val !== undefined && val !== '') {
                    textValues.push({ r, c, val: strVal });
                }
            }
        }

        const sorted = [...values].sort((a, b) => a.val - b.val);
        const numVals = values.map(v => v.val);

        this._rangeDataCache = {
            values,
            textValues,
            sorted,
            cellMap,
            min: sorted.length > 0 ? sorted[0].val : 0,
            max: sorted.length > 0 ? sorted[sorted.length - 1].val : 0,
            avg: numVals.length > 0 ? numVals.reduce((a, b) => a + b, 0) / numVals.length : 0,
            count: values.length,
            totalCount: textValues.length
        };

        return this._rangeDataCache;
    }

    /** @param {string|number} formula @param {number} r @param {number} c @returns {*} */
    resolveFormula(formula, r, c) {
        if (formula === null || formula === undefined) return null;
        if (typeof formula === 'number') return formula;

        // 字符串公式
        const f = String(formula).trim();

        // 纯数字字符串
        const num = parseFloat(f);
        if (!isNaN(num) && String(num) === f) return num;

        // 公式计算（支持跨表引用）
        if (this._SN.Formula) {
            try {
                return this._SN.Formula.calcFormula(f, {
                    row: { rIndex: r },
                    cIndex: c
                });
            } catch (e) {
                console.warn('CF formula error:', e);
                return null;
            }
        }

        return f;
    }

    _clearCache() {
        this._evalCache.clear();
        this._rangeDataCache = null;
    }

    /** @param {Object} cf @returns {Object} */
    _toXmlObject(cf) {
        const node = {
            '_$type': this.type,
            '_$priority': String(this.priority)
        };

        if (this.stopIfTrue) {
            node['_$stopIfTrue'] = '1';
        }

        if (this.dxf) {
            node['_$dxfId'] = String(cf.getDxfId(this.dxf));
        }

        // subclass override
        this._addXmlAttributes(node, cf);

        return node;
    }

    _addXmlAttributes(node, cf) {
    }
}

export default CFRule;
