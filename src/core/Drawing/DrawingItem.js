import { chartOptionApplyTemplate, chartXmlToEChartOption, refOptionParse } from "./chart.js";
import { _initImage } from "./image.js";
import { parseShapeFromXml, buildShapeXml } from "./shape.js";
import { CONNECTOR_TYPES } from "./constants.js";

/** @class */
export default class DrawingItem {
    /** @param {Object} config @param {Sheet} sheet */
    constructor(config, sheet) {

        /** @type {Sheet} */
        this.sheet = sheet;
        this._type = config.type;
        /** @type {string} */
        this.id = "rId" + ++sheet._idCount;
        /** @type {boolean} */
        this.inView = false
        /** @type {string} */
        this.anchorType = config.anchorType ?? 'twoCell';
        /** @type {Object|null} */
        this.xmlDetail = config.xmlDetail ?? null;

        // default start cell: top-left of viewport
        const defaultStartCell = {
            r: sheet.vi.rowArr[0] ?? 0,
            c: sheet.vi.colArr[0] ?? 0
        };
        const startCell = config.startCell;
        this._startCell = typeof startCell == 'string'
            ? sheet.Utils.cellStrToNum(startCell)
            : startCell ?? defaultStartCell;

        this._offsetX = config.offsetX ?? 0;
        this._offsetY = config.offsetY ?? 0;

        // image type: no default size, resolved by _initImage
        if (config.type === 'image') {
            this._width = config.width ?? null;
            this._height = config.height ?? null;
        } else if (config.type === 'shape') {
            this._width = config.width ?? 100;
            this._height = config.height ?? 100;
        } else if (config.type === 'slicer') {
            this._width = config.width ?? 180;
            this._height = config.height ?? 220;
        } else {
            this._width = config.width ?? 460;
            this._height = config.height ?? 260;
        }
        this._chartOption = (config.chartOption && chartOptionApplyTemplate(config.chartOption, sheet?.SN)) ?? null;
        this._area = null;
        this._renderOption = null;
        this._drawCache = null;
        this._updRender = false;

        this._imageBase64 = config.imageBase64 ?? null;
        this._imageSuffix = null;
        this._rotation = config.rotation ?? 0;

        // twoCell end cell reference
        this._endCell = null;
        this._endOffsetX = 0;
        this._endOffsetY = 0;

        // shape related
        /** @type {string} */
        this.shapeType = config.shapeType ?? 'rect';
        /** @type {boolean} */
        this.isConnector = CONNECTOR_TYPES.includes(this.shapeType);
        /** @type {boolean} */
        this.isTextBox = !this.isConnector && !!(config.isTextBox ?? config.textBox);

        // unified style system
        const defaultShapeStyle = this.isConnector ? {
            fill: 'none',
            stroke: '#000000',
            strokeWidth: 1,
            textColor: '#000000',
            fontSize: 14,
            startArrow: 'none',
            endArrow: 'none'
        } : this.isTextBox ? {
            fill: '#FFFFFF',
            stroke: '#7F7F7F',
            strokeWidth: 0.75,
            textColor: '#000000',
            fontSize: 11,
            textAlign: 'l',
            textDirection: 'horizontal',
            startArrow: 'none',
            endArrow: 'none'
        } : {
            fill: '#4472C4',
            stroke: '#000000',
            strokeWidth: 1,
            textColor: '#000000',
            fontSize: 14,
            textDirection: 'horizontal',
            startArrow: 'none',
            endArrow: 'none'
        };
        /** @type {Object} textDirection:horizontal|vertical */
        this.shapeStyle = { ...defaultShapeStyle, ...(config.shapeStyle ?? {}) };
        if (!this.shapeStyle.textDirection) this.shapeStyle.textDirection = 'horizontal';
        if (this.isTextBox && !this.shapeStyle.textAlign) this.shapeStyle.textAlign = 'l';

        if (config.arrows) {
            switch (config.arrows) {
                case 'none':
                    this.shapeStyle.startArrow = 'none';
                    this.shapeStyle.endArrow = 'none';
                    break;
                case 'end':
                    this.shapeStyle.startArrow = 'none';
                    this.shapeStyle.endArrow = 'arrow';
                    break;
                case 'both':
                    this.shapeStyle.startArrow = 'arrow';
                    this.shapeStyle.endArrow = 'arrow';
                    break;
            }
        }

        /** @type {string} */
        this.shapeText = config.shapeText ?? '';

        // parse shape config from XML
        if (this.type === 'shape' && this.xmlDetail) {
            const parsed = parseShapeFromXml(this.xmlDetail, sheet.SN, this.isConnector);
            if (parsed) {
                this.shapeType = parsed.shapeType;
                this.isConnector = CONNECTOR_TYPES.includes(parsed.shapeType);
                this.isTextBox = this.isConnector ? false : (parsed.isTextBox ?? this.isTextBox);
                this.shapeStyle = { ...this.shapeStyle, ...parsed.shapeStyle };
                if (!this.shapeStyle.textDirection) this.shapeStyle.textDirection = 'horizontal';
                if (this.isTextBox && !this.shapeStyle.textAlign) this.shapeStyle.textAlign = 'l';
                this.shapeText = parsed.shapeText;
            }
        }

        this._initPromise = null;

        // auto init for image type
        if (this.type === 'image' && this._imageBase64) {
            this._initPromise = this._initImage();
        }
    }

    /** @type {'chart'|'image'|'shape'|'slicer'} */
    get type() {
        return this._type;
    }

    /** @param {'up'|'down'|'top'|'bottom'} type */
    updIndex(type) {
        const drawing = this.sheet.Drawing;
        const drawings = drawing.getAll();
        const index = drawings.indexOf(this);
        if (index === -1) return;

        switch (type) {
            case 'up':
                if (index < drawings.length - 1) {
                    [drawings[index], drawings[index + 1]] = [drawings[index + 1], drawings[index]];
                }
                break;
            case 'down':
                if (index > 0) {
                    [drawings[index], drawings[index - 1]] = [drawings[index - 1], drawings[index]];
                }
                break;
            case 'top':
                drawings.splice(index, 1);
                drawings.push(this);
                break;
            case 'bottom':
                drawings.splice(index, 1);
                drawings.unshift(this);
                break;
        }
    }


    /** @returns {Object} */
    get imageBytes() {
        const base64Data = this.imageBase64.split(',')[1];

        // Base64 → 二进制字符串
        const binaryString = atob(base64Data);

        // 二进制字符串 → Uint8Array
        const len = binaryString.length;
        const uint8Array = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            uint8Array[i] = binaryString.charCodeAt(i);
        }
        return {
            _data: uint8Array
        }
    }

    /** @type {string|null} */
    get imageSuffix() {
        if (this._imageSuffix) return this._imageSuffix
        return this._imageSuffix = this.imageBase64.match(/^data:image\/([^;]+);base64,/)?.[1] || null
    }

    /** @type {string} */
    get imageBase64() {
        if (this._imageBase64) return this._imageBase64
        if (this.xmlDetail) {
            const ext = this.xmlDetail.name.split('.').at(-1)
            this._imageSuffix = ext
            return this._imageBase64 = `data:image/${ext};base64,` + btoa(this.xmlDetail.asBinary())
        } else {
            return ""
        }
    }

    /** @type {string} */
    set imageBase64(base64) {
        if (!base64?.startsWith('data:image')) base64 = `data:image/png;base64,` + base64
        this._imageBase64 = base64
        this._imageSuffix = null
        this.drawCache = null
    }

    /** @type {boolean} */
    get updRender() {
        return this._updRender
    }

    set updRender(bool) {
        this._updRender = bool
        if (bool) {
            this._renderOption = null
            this.drawCache = null
        }
    }

    /** @type {CellNum} */
    get startCell() {
        return this._startCell;
    }

    set startCell(val) {
        this._startCell = val;
        this._area = null
    }

    /** @type {HTMLElement|null} */
    get drawCache() {
        return this._drawCache;
    }

    set drawCache(dom) {
        if (this._drawCache) {
            // Canvas 元素清理显存
            if (this._drawCache instanceof HTMLCanvasElement) {
                this._drawCache.width = 0;
                this._drawCache.height = 0;
            }
            // DOM 元素移除
            this._drawCache?.remove?.();
        }
        this._drawCache = dom
    }

    /** @type {Object} */
    get area() {
        // twoCell 模式：从起止单元格引用直接构建（不缓存，实时响应行列变化）
        if (this.anchorType === 'twoCell' && this._endCell) {
            return {
                s: { r: this._startCell.r, c: this._startCell.c, offsetX: this._offsetX, offsetY: this._offsetY },
                e: { r: this._endCell.r, c: this._endCell.c, offsetX: this._endOffsetX, offsetY: this._endOffsetY }
            };
        }

        if (this._area) return this._area

        // compute end position from width/height (hidden rows/cols occupy 0px on screen)
        let remW = this.width + (this._visColW(this.startCell.c) ? this.offsetX : 0);
        let endC = this.startCell.c;
        while (remW > this._visColW(endC)) {
            remW -= this._visColW(endC);
            endC++;
        }
        const endOffsetX = remW;

        let remH = this.height + (this._visRowH(this.startCell.r) ? this.offsetY : 0);
        let endR = this.startCell.r;
        while (remH > this._visRowH(endR)) {
            remH -= this._visRowH(endR);
            endR++;
        }
        const endOffsetY = remH;

        return this._area = {
            s: {
                "r": this.startCell.r,
                "c": this.startCell.c,
                "offsetX": this.offsetX,
                "offsetY": this.offsetY
            },
            e: {
                "r": endR,
                "c": endC,
                "offsetX": endOffsetX,
                "offsetY": endOffsetY
            }
        }
    }

    /** @type {number} */
    get offsetX() {
        return this._offsetX;
    }

    set offsetX(num) {
        if (num > this.sheet.getCol(this.startCell.c).width) throw new Error("offsetX exceeds start cell width")
        this._offsetX = num;
        this._area = null
    }

    /** @type {number} */
    get offsetY() {
        return this._offsetY;
    }

    set offsetY(num) {
        if (num > this.sheet.getRow(this.startCell.r).height) throw new Error("offsetY exceeds start cell height")
        this._offsetY = num;
        this._area = null
    }

    /** @type {number} */
    get rotation() {
        return this._rotation;
    }

    set rotation(deg) {
        this._rotation = ((deg % 360) + 360) % 360;
    }

    /** @type {number} */
    get width() {
        // twoCell 模式：从单元格网格实时计算像素宽度（隐藏列按 0 计）
        if (this.anchorType === 'twoCell' && this._endCell) {
            let w = this._visColW(this._startCell.c) ? -this._offsetX : 0;
            for (let c = this._startCell.c; c < this._endCell.c; c++) {
                w += this._visColW(c);
            }
            return w + (this._visColW(this._endCell.c) ? this._endOffsetX : 0);
        }
        // 惰性计算：如果_width是null且有_area，从area反推
        if (this._width === null && this._area) {
            let width = 0;
            for (let c = this._area.s.c; c <= this._area.e.c; c++) {
                width += this._visColW(c);
            }
            const sOff = this._visColW(this._area.s.c) ? this._area.s.offsetX : 0;
            const eTail = this._visColW(this._area.e.c) ? this._visColW(this._area.e.c) - this._area.e.offsetX : 0;
            this._width = width - sOff - eTail;
        }
        return this._width;
    }

    set width(num) {
        this._width = num;
        this._area = null;
        if (this.anchorType === 'twoCell') this._syncEndCell();
    }

    /** @type {number} */
    get height() {
        // twoCell 模式：从单元格网格实时计算像素高度（隐藏行按 0 计）
        if (this.anchorType === 'twoCell' && this._endCell) {
            let h = this._visRowH(this._startCell.r) ? -this._offsetY : 0;
            for (let r = this._startCell.r; r < this._endCell.r; r++) {
                h += this._visRowH(r);
            }
            return h + (this._visRowH(this._endCell.r) ? this._endOffsetY : 0);
        }
        // 惰性计算：如果_height是null且有_area，从area反推
        if (this._height === null && this._area) {
            let height = 0;
            for (let r = this._area.s.r; r <= this._area.e.r; r++) {
                height += this._visRowH(r);
            }
            const sOff = this._visRowH(this._area.s.r) ? this._area.s.offsetY : 0;
            const eTail = this._visRowH(this._area.e.r) ? this._visRowH(this._area.e.r) - this._area.e.offsetY : 0;
            this._height = height - sOff - eTail;
        }
        return this._height;
    }

    set height(num) {
        this._height = num;
        this._area = null;
        if (this.anchorType === 'twoCell') this._syncEndCell();
    }

    /** @type {Object|null} */
    get chartOption() {
        if (!this._chartOption) this._chartOption = chartXmlToEChartOption(this.xmlDetail, this.sheet.SN); // 如果chartOption为null，就尝试去xmlDetail解析
        return this._chartOption;
    }

    set chartOption(obj) {
        if (this.type == 'chart') obj = chartOptionApplyTemplate(obj, this.sheet?.SN);
        this._chartOption = obj
        this._renderOption = null
        this._hasChartRefs = undefined
        this.drawCache = null
    }

    /** @type {Object|null} */
    get renderOption() {
        if (this._renderOption) return this._renderOption;
        this._hasChartReferences();
        return this._renderOption = refOptionParse(this.chartOption, this.sheet.SN);
    }

    /** @returns {boolean} */
    _hasChartReferences() {
        if (this.type !== 'chart') return false;
        if (this._hasChartRefs !== undefined) return this._hasChartRefs;

        const refRegex = /(?:'[^']+'|[^'!]+)!\$?[A-Za-z]{1,3}\$?\d+/;
        this._hasChartRefs = refRegex.test(JSON.stringify(this.chartOption));
        return this._hasChartRefs;
    }

    _clearRenderCache() {
        this._renderOption = null;
        this.drawCache = null;
    }

    // ==================== Anchor behavior ====================

    /** on-screen col width: hidden cols occupy 0px @param {number} c @returns {number} */
    _visColW(c) {
        const col = this.sheet.getCol(c);
        return col.hidden ? 0 : col.width;
    }

    /** on-screen row height: hidden/filtered rows occupy 0px @param {number} r @returns {number} */
    _visRowH(r) {
        const row = this.sheet.getRow(r);
        return row.hidden ? 0 : row.height;
    }

    /** sync _endCell from pixel size (twoCell mode); pixel size is on-screen size, so hidden rows/cols count 0 */
    _syncEndCell() {
        if (this._width === null || this._height === null) return;

        let remW = this._width + (this._visColW(this._startCell.c) ? this._offsetX : 0);
        let endC = this._startCell.c;
        while (remW > this._visColW(endC)) {
            remW -= this._visColW(endC);
            endC++;
        }

        let remH = this._height + (this._visRowH(this._startCell.r) ? this._offsetY : 0);
        let endR = this._startCell.r;
        while (remH > this._visRowH(endR)) {
            remH -= this._visRowH(endR);
            endR++;
        }

        this._endCell = { r: endR, c: endC };
        this._endOffsetX = remW;
        this._endOffsetY = remH;
    }

    /** @param {number} r @param {number} n */
    _adjustForRowInsert(r, n) {
        if (this.anchorType === 'absolute') return;

        // oneCell / twoCell：起始行在插入位置及之后则下移
        if (this._startCell.r >= r) {
            this._startCell = { ...this._startCell, r: this._startCell.r + n };
        }
        // twoCell：结束行在插入位置及之后则下移（若在图纸中间则图纸变高）
        if (this.anchorType === 'twoCell' && this._endCell && this._endCell.r >= r) {
            this._endCell = { ...this._endCell, r: this._endCell.r + n };
        }
        this._area = null;
    }

    /** @param {number} r @param {number} n */
    _adjustForRowDelete(r, n) {
        if (this.anchorType === 'absolute') return;

        const delEnd = r + n; // 删除范围 [r, delEnd)

        // 调整起始行
        if (this._startCell.r >= delEnd) {
            this._startCell = { ...this._startCell, r: this._startCell.r - n };
        } else if (this._startCell.r >= r) {
            this._startCell = { ...this._startCell, r: r };
            this._offsetY = 0;
        }

        // twoCell：调整结束行
        if (this.anchorType === 'twoCell' && this._endCell) {
            if (this._endCell.r >= delEnd) {
                this._endCell = { ...this._endCell, r: this._endCell.r - n };
            } else if (this._endCell.r >= r) {
                this._endCell = { ...this._endCell, r: r };
                this._endOffsetY = 0;
            }
        }
        this._area = null;
    }

    /** @param {number} c @param {number} n */
    _adjustForColInsert(c, n) {
        if (this.anchorType === 'absolute') return;

        if (this._startCell.c >= c) {
            this._startCell = { ...this._startCell, c: this._startCell.c + n };
        }
        if (this.anchorType === 'twoCell' && this._endCell && this._endCell.c >= c) {
            this._endCell = { ...this._endCell, c: this._endCell.c + n };
        }
        this._area = null;
    }

    /** @param {number} c @param {number} n */
    _adjustForColDelete(c, n) {
        if (this.anchorType === 'absolute') return;

        const delEnd = c + n;

        if (this._startCell.c >= delEnd) {
            this._startCell = { ...this._startCell, c: this._startCell.c - n };
        } else if (this._startCell.c >= c) {
            this._startCell = { ...this._startCell, c: c };
            this._offsetX = 0;
        }

        if (this.anchorType === 'twoCell' && this._endCell) {
            if (this._endCell.c >= delEnd) {
                this._endCell = { ...this._endCell, c: this._endCell.c - n };
            } else if (this._endCell.c >= c) {
                this._endCell = { ...this._endCell, c: c };
                this._endOffsetX = 0;
            }
        }
        this._area = null;
    }

}

// 使用原型链挂载图片模块方法
Object.assign(DrawingItem.prototype, { _initImage });

// 使用原型链挂载形状模块方法
Object.assign(DrawingItem.prototype, { buildShapeXml });
