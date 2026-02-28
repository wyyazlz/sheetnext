import Cell from '../Cell/Cell.js';

// xmlObj可不传，当赋值或设置此列东西时才去初始化一个

/**
 * 行对象
 * @title ↔️ 行操作
 * @class
 */
export default class Row {
    #height;
    #hidden;
    #filterHidden;
    #outlineLevel;
    #collapsed;
    #style = null;  // 行样式

    /**
     * @param {Object} xmlObj - 行 XML 数据
     * @param {Sheet} sheet - 所属工作表
     * @param {number} rIndex - 行索引
     */
    constructor(xmlObj, sheet, rIndex) {
        /**
         * 所属工作表
         * @type {Sheet}
         */
        this.sheet = sheet
        /**
         * 行索引
         * @type {number}
         */
        this.rIndex = rIndex
        /**
         * 行内单元格数组
         * @type {Array<Cell>}
         */
        this.cells = [];
        this._SN = sheet.SN
        this._xmlObj = xmlObj ?? {}
        this.#height = this._xmlObj['_$ht'] ? Math.round(this._xmlObj['_$ht'] * 1.333) : this.sheet.defaultRowHeight
        this.#hidden = this._xmlObj['_$hidden'] == "1"
        this.#filterHidden = false;
        this.#outlineLevel = Math.max(0, Math.min(7, Math.floor(Number(this._xmlObj['_$outlineLevel']) || 0)));
        this.#collapsed = this._xmlObj['_$collapsed'] == '1';
        // 解析行样式
        if (this._xmlObj['_$s'] && this._xmlObj['_$customFormat'] == '1') {
            this.#style = this._SN.Xml._xGetStyle(this._xmlObj['_$s']);
            if (this.#style && Object.keys(this.#style).length > 0) {
                this.sheet._styledRows.add(rIndex);
            }
        }
        this.init();
    }

    /**
     * 初始化行内单元格
     * @returns {void}
     */
    init() {
        if (!this._xmlObj.c) {
            return
        } else if (!Array.isArray(this._xmlObj.c)) {
            this._xmlObj.c = [this._xmlObj.c];
        }
        this._xmlObj.c.forEach(item => {
            const cIndex = this.sheet.Utils.charToNum(item['_$r'].replace(/\d/g, ''));
            this.cells[cIndex] = new Cell(item, this, cIndex);
        })
    }

    // ========== 内部属性（自动清缓存，不触发 UndoRedo）==========
    get _height() { return this.#height; }
    set _height(v) {
        this.#height = v;
        this.sheet._totalHeightCache = undefined;
    }

    get _hidden() { return this.#hidden; }
    set _hidden(v) {
        if (this.#hidden === v) return;
        this.#hidden = v;
        this.sheet._totalHeightCache = undefined;
        this.sheet._rowVisibilityVersion = (this.sheet._rowVisibilityVersion || 0) + 1;
    }

    get _filterHidden() { return this.#filterHidden; }
    set _filterHidden(v) {
        const next = !!v;
        if (this.#filterHidden === next) return;
        this.#filterHidden = next;
        this.sheet._totalHeightCache = undefined;
        this.sheet._rowVisibilityVersion = (this.sheet._rowVisibilityVersion || 0) + 1;
    }

    get _outlineLevel() { return this.#outlineLevel; }
    set _outlineLevel(v) {
        this.#outlineLevel = Math.max(0, Math.min(7, Math.floor(Number(v) || 0)));
    }

    get _collapsed() { return this.#collapsed; }
    set _collapsed(v) {
        this.#collapsed = !!v;
    }

    // ========== 公开属性（触发 UndoRedo）==========
    /**
     * 行高（像素）
     * @type {number}
     */
    get height() { return this.#height; }
    set height(num) {
        const old = this.#height;
        if (old === num) return;
        if (this._SN.Event.hasListeners('beforeRowResize')) {
            const beforeEvent = this._SN.Event.emit('beforeRowResize', {
                sheet: this.sheet,
                row: this.rIndex,
                oldHeight: old,
                newHeight: num
            });
            if (beforeEvent.canceled) return;
        }
        return this._SN._withOperation('rowResize', {
            sheet: this.sheet,
            row: this.rIndex,
            oldHeight: old,
            newHeight: num
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._height = old; },
                redo: () => { this._height = num; }
            });
            this._height = num;
            this._SN._recordChange({
                type: 'row',
                action: 'resize',
                sheet: this.sheet,
                row: this.rIndex,
                oldHeight: old,
                newHeight: num
            });
            if (this._SN.Event.hasListeners('afterRowResize')) {
                this._SN.Event.emit('afterRowResize', {
                    sheet: this.sheet,
                    row: this.rIndex,
                    oldHeight: old,
                    newHeight: num
                });
            }
        });
    }

    /**
     * 是否隐藏
     * @type {boolean}
     */
    get hidden() { return this.#hidden || this.#filterHidden; }
    set hidden(bol) {
        const old = this.#hidden;
        const next = !!bol;
        if (old === next) return;
        const eventBase = next ? 'RowHide' : 'RowUnhide';
        if (this._SN.Event.hasListeners(`before${eventBase}`)) {
            const beforeEvent = this._SN.Event.emit(`before${eventBase}`, {
                sheet: this.sheet,
                row: this.rIndex,
                oldHidden: old,
                newHidden: next
            });
            if (beforeEvent.canceled) return;
        }
        return this._SN._withOperation('rowHidden', {
            sheet: this.sheet,
            row: this.rIndex,
            oldHidden: old,
            newHidden: next
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._hidden = old; },
                redo: () => { this._hidden = next; }
            });
            this._hidden = next;
            this._SN._recordChange({
                type: 'row',
                action: next ? 'hide' : 'unhide',
                sheet: this.sheet,
                row: this.rIndex,
                oldHidden: old,
                newHidden: next
            });
            if (this._SN.Event.hasListeners(`after${eventBase}`)) {
                this._SN.Event.emit(`after${eventBase}`, {
                    sheet: this.sheet,
                    row: this.rIndex,
                    oldHidden: old,
                    newHidden: next
                });
            }
        });
    }

    /**
     * 大纲层级（0-7）
     * @type {number}
     */
    get outlineLevel() { return this.#outlineLevel; }
    set outlineLevel(num) {
        const old = this.#outlineLevel;
        const next = Math.max(0, Math.min(7, Math.floor(Number(num) || 0)));
        if (old === next) return;
        return this._SN._withOperation('rowOutlineLevel', {
            sheet: this.sheet,
            row: this.rIndex,
            oldOutlineLevel: old,
            newOutlineLevel: next
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._outlineLevel = old; },
                redo: () => { this._outlineLevel = next; }
            });
            this._outlineLevel = next;
            this._SN._recordChange({
                type: 'row',
                action: 'outlineLevel',
                sheet: this.sheet,
                row: this.rIndex,
                oldOutlineLevel: old,
                newOutlineLevel: next
            });
        });
    }

    /**
     * 大纲折叠标记
     * @type {boolean}
     */
    get collapsed() { return this.#collapsed; }
    set collapsed(val) {
        const old = this.#collapsed;
        const next = !!val;
        if (old === next) return;
        return this._SN._withOperation('rowCollapsed', {
            sheet: this.sheet,
            row: this.rIndex,
            oldCollapsed: old,
            newCollapsed: next
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._collapsed = old; },
                redo: () => { this._collapsed = next; }
            });
            this._collapsed = next;
            this._SN._recordChange({
                type: 'row',
                action: 'collapsed',
                sheet: this.sheet,
                row: this.rIndex,
                oldCollapsed: old,
                newCollapsed: next
            });
        });
    }

    /**
     * 是否需要写入 XML
     * @type {boolean}     */
    get _needBuild() {
        if (this.#height !== this.sheet.defaultRowHeight || this.hidden || this.#style || this.#outlineLevel > 0 || this.#collapsed) return true
        if (this.cells.some(c => c?._needBuild)) return true
        return false;
    }

    // ========== 行样式 ==========
    get _rowStyle() { return this.#style; }
    set _rowStyle(v) {
        this.#style = v;
        if (v) this.sheet._styledRows.add(this.rIndex);
        else this.sheet._styledRows.delete(this.rIndex);
    }

    /**
     * 行样式
     * @type {Object}
     */
    get style() { return this.#style || {}; }
    set style(obj) {
        const old = this.#style;
        const oldInSet = this.sheet._styledRows.has(this.rIndex);
        this._rowStyle = obj ? { ...obj } : null;
        this._SN.UndoRedo.add({
            undo: () => {
                this.#style = old;
                if (oldInSet) this.sheet._styledRows.add(this.rIndex);
                else this.sheet._styledRows.delete(this.rIndex);
            },
            redo: () => {
                this._rowStyle = obj ? { ...obj } : null;
            }
        });
    }

    /**
     * 字体样式
     * @type {Object}
     */
    get font() { return this.#style?.font || {}; }
    set font(obj) {
        const nextFont = obj ? { ...this.#style?.font, ...obj } : undefined;
        if (nextFont && Object.prototype.hasOwnProperty.call(obj, 'color')) {
            delete nextFont.colorTheme;
            delete nextFont.colorTint;
            delete nextFont.colorIndexed;
            delete nextFont.colorAuto;
        }
        const newStyle = { ...this.#style, font: nextFont };
        if (!newStyle.font) delete newStyle.font;
        this.style = Object.keys(newStyle).length ? newStyle : null;
    }

    /**
     * 填充样式
     * @type {Object}
     */
    get fill() { return this.#style?.fill || {}; }
    set fill(obj) {
        const newStyle = { ...this.#style, fill: obj ? { type: 'pattern', pattern: 'solid', ...this.#style?.fill, ...obj } : undefined };
        if (!newStyle.fill) delete newStyle.fill;
        this.style = Object.keys(newStyle).length ? newStyle : null;
    }

    /**
     * 对齐样式
     * @type {Object}
     */
    get alignment() { return this.#style?.alignment || {}; }
    set alignment(obj) {
        const newStyle = { ...this.#style, alignment: obj ? { ...this.#style?.alignment, ...obj } : undefined };
        if (!newStyle.alignment) delete newStyle.alignment;
        this.style = Object.keys(newStyle).length ? newStyle : null;
    }

    /**
     * 边框样式
     * @type {Object}
     */
    get border() { return this.#style?.border || {}; }
    set border(obj) {
        let nextBorder;
        if (obj) {
            nextBorder = { ...(this.#style?.border || {}) };
            Object.keys(obj).forEach((name) => {
                const value = obj[name];
                if (['top', 'right', 'bottom', 'left', 'diagonal'].includes(name)) {
                    if (value && typeof value === 'object') {
                        nextBorder[name] = { color: '#000000', style: 'thin', ...(nextBorder[name] || {}), ...value };
                    } else {
                        delete nextBorder[name];
                    }
                    return;
                }
                if (name === 'diagonalUp' || name === 'diagonalDown') {
                    if (value === undefined || value === null) delete nextBorder[name];
                    else nextBorder[name] = !!value;
                    return;
                }
                if (value === undefined || value === null) delete nextBorder[name];
                else nextBorder[name] = value;
            });
            if (!Object.keys(nextBorder).length) nextBorder = undefined;
        }
        const newStyle = { ...this.#style, border: nextBorder };
        if (!newStyle.border) delete newStyle.border;
        this.style = Object.keys(newStyle).length ? newStyle : null;
    }

    /**
     * 数字格式
     * @type {string|undefined}
     */
    get numFmt() { return this.#style?.numFmt; }
    set numFmt(val) {
        const newStyle = { ...this.#style, numFmt: val || undefined };
        if (!newStyle.numFmt) delete newStyle.numFmt;
        this.style = Object.keys(newStyle).length ? newStyle : null;
    }

}
