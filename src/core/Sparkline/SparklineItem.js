/** @class */
export default class SparklineItem {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        this.sheet = sheet;
        /** @type {number} */
        this.r = config.r;
        /** @type {number} */
        this.c = config.c;
        /** @type {string} */
        this.sqref = config.sqref;
        /** @type {string} */
        this.formula = config.formula;
        this.group = config.group;
        this._dataCache = null;
        this._renderCache = null;
    }

    /** @type {string} */
    get key() {
        return `${this.r}:${this.c}`;
    }

    /** @type {'line'|'column'|'stacked'} */
    get type() { return this.group.type; }
    set type(v) { this.group.type = v; }

    /** @type {Object} */
    get colors() { return this.group.colors; }
    set colors(v) { this.group.colors = v; }

    /** @returns {Array<number|null>} */
    getData() {
        if (this._dataCache) return this._dataCache;

        const formula = this.formula;
        // 解析公式，如 "Sheet1!B2:D2"
        const match = formula.match(/^(.+?)!(.+)$/);
        if (!match) return [];

        const [, sheetName, rangeStr] = match;
        const cleanSheetName = sheetName.replace(/^'|'$/g, '');
        const targetSheet = this.sheet.SN.getSheet(cleanSheetName) || this.sheet;

        // 解析范围
        const range = targetSheet.rangeStrToNum(rangeStr);
        const data = [];

        // 遍历范围获取数值
        targetSheet.eachCells(range, (r, c) => {
            const cell = targetSheet.getCell(r, c);
            const val = cell.calcVal ?? cell.editVal;
            if (val !== null && val !== undefined && val !== '') {
                const num = parseFloat(val);
                data.push(isNaN(num) ? null : num);
            } else {
                data.push(null);
            }
        });

        this._dataCache = data;
        return data;
    }

    _clearCache() {
        this._dataCache = null;
        this._renderCache = null;
    }

    /** @returns {Object} */
    _toXmlObject() {
        return {
            'xm:f': this.formula,
            'xm:sqref': this.sqref
        };
    }
}
