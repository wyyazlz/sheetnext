/** @class */
export default class CommentItem {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {
        this.sheet = sheet;
        /** @type {string} */
        this.id = config.id ?? `comment_${Date.now()}_${Math.random()}`;
        /** @type {string} */
        this.cellRef = config.cellRef;
        /** @type {string} */
        this.author = config.author ?? '未知作者';
        /** @type {string} */
        this.text = config.text ?? '';
        /** @type {boolean} */
        this.visible = config.visible ?? false;
        /** @type {number} */
        this.width = config.width ?? 144;
        /** @type {number} */
        this.height = config.height ?? 72;
        /** @type {number} */
        this.marginLeft = config.marginLeft ?? 107;
        /** @type {number} */
        this.marginTop = config.marginTop ?? 7;

        this._position = null;
        this._cellPos = null;
    }

    /** @type {CellNum} */
    get cellPos() {
        if (!this._cellPos) {
            this._cellPos = this.sheet.SN.Utils.cellStrToNum(this.cellRef);
        }
        return this._cellPos;
    }

    /** @returns {Object|null} */
    _getPosition() {
        if (this._position) return this._position;

        const { r, c } = this.cellPos;
        const cellInfo = this.sheet.getCellInViewInfo(r, c);

        if (!cellInfo) return null;

        // 将pt转换为px（1pt ≈ 1.33px）
        const pxPerPt = 1.33;

        this._position = {
            x: cellInfo.x + this.marginLeft * pxPerPt,
            y: cellInfo.y + this.marginTop * pxPerPt,
            width: this.width * pxPerPt,
            height: this.height * pxPerPt
        };

        return this._position;
    }

    _clearCache() {
        this._position = null;
    }

    /** @returns {Object} */
    _toXmlObject() {
        return {
            ref: this.cellRef,
            authorId: 0, // 简化处理，统一为0
            text: {
                t: this.text
            }
        };
    }

    /** @returns {Object} */
    _toVmlObject() {
        return {
            visible: this.visible,
            width: this.width,
            height: this.height,
            marginLeft: this.marginLeft,
            marginTop: this.marginTop
        };
    }
}
