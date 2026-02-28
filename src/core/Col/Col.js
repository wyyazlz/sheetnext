/**
 * 列对象
 * @title ↕️ 列操作
 * @class
 */
export default class Col {
    #width;
    #hidden;
    #outlineLevel;
    #collapsed;
    #style = null;  // 列样式

    /**
     * @param {Object} xmlObj - 列 XML 数据
     * @param {Sheet} sheet - 所属工作表
     * @param {number} cIndex - 列索引
     */
    constructor(xmlObj, sheet, cIndex) {
        /**
         * 所属工作表
         * @type {Sheet}
         */
        this.sheet = sheet
        /**
         * 列索引
         * @type {number}
         */
        this.cIndex = cIndex
        this._SN = sheet.SN
        this._xmlObj = xmlObj ?? {}
        this.#hidden = this._xmlObj['_$hidden'] == '1'
        this.#width = this._xmlObj['_$width'] ? Math.round(this._xmlObj['_$width'] * 7.5) : this.sheet.defaultColWidth
        this.#outlineLevel = Math.max(0, Math.min(7, Math.floor(Number(this._xmlObj['_$outlineLevel']) || 0)));
        this.#collapsed = this._xmlObj['_$collapsed'] == '1';
        // 解析列样式
        if (this._xmlObj['_$style']) {
            this.#style = this._SN.Xml._xGetStyle(this._xmlObj['_$style']);
            if (this.#style && Object.keys(this.#style).length > 0) {
                this.sheet._styledCols.add(cIndex);
            }
        }
    }

    /**
     * 获取列内所有单元格
     * @type {Array<Cell>}     */
    get cells() {
        return this.sheet.rows.map(row => row.cells[this.cIndex]);
    }

    // ========== 内部属性（自动清缓存，不触发 UndoRedo）==========
    get _width() { return this.#width; }
    set _width(v) {
        this.#width = v;
        this.sheet._totalWidthCache = undefined;
    }

    get _hidden() { return this.#hidden; }
    set _hidden(v) {
        this.#hidden = v;
        this.sheet._totalWidthCache = undefined;
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
     * 列宽（像素）
     * @type {number}
     */
    get width() { return this.#width; }
    set width(num) {
        const old = this.#width;
        if (old === num) return;
        if (this._SN.Event.hasListeners('beforeColResize')) {
            const beforeEvent = this._SN.Event.emit('beforeColResize', {
                sheet: this.sheet,
                col: this.cIndex,
                oldWidth: old,
                newWidth: num
            });
            if (beforeEvent.canceled) return;
        }
        return this._SN._withOperation('colResize', {
            sheet: this.sheet,
            col: this.cIndex,
            oldWidth: old,
            newWidth: num
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._width = old; },
                redo: () => { this._width = num; }
            });
            this._width = num;
            this._SN._recordChange({
                type: 'col',
                action: 'resize',
                sheet: this.sheet,
                col: this.cIndex,
                oldWidth: old,
                newWidth: num
            });
            if (this._SN.Event.hasListeners('afterColResize')) {
                this._SN.Event.emit('afterColResize', {
                    sheet: this.sheet,
                    col: this.cIndex,
                    oldWidth: old,
                    newWidth: num
                });
            }
        });
    }

    /**
     * 是否隐藏
     * @type {boolean}
     */
    get hidden() { return this.#hidden; }
    set hidden(bol) {
        const old = this.#hidden;
        const next = !!bol;
        if (old === next) return;
        const eventBase = next ? 'ColHide' : 'ColUnhide';
        if (this._SN.Event.hasListeners(`before${eventBase}`)) {
            const beforeEvent = this._SN.Event.emit(`before${eventBase}`, {
                sheet: this.sheet,
                col: this.cIndex,
                oldHidden: old,
                newHidden: next
            });
            if (beforeEvent.canceled) return;
        }
        return this._SN._withOperation('colHidden', {
            sheet: this.sheet,
            col: this.cIndex,
            oldHidden: old,
            newHidden: next
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._hidden = old; },
                redo: () => { this._hidden = next; }
            });
            this._hidden = next;
            this._SN._recordChange({
                type: 'col',
                action: next ? 'hide' : 'unhide',
                sheet: this.sheet,
                col: this.cIndex,
                oldHidden: old,
                newHidden: next
            });
            if (this._SN.Event.hasListeners(`after${eventBase}`)) {
                this._SN.Event.emit(`after${eventBase}`, {
                    sheet: this.sheet,
                    col: this.cIndex,
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
        return this._SN._withOperation('colOutlineLevel', {
            sheet: this.sheet,
            col: this.cIndex,
            oldOutlineLevel: old,
            newOutlineLevel: next
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._outlineLevel = old; },
                redo: () => { this._outlineLevel = next; }
            });
            this._outlineLevel = next;
            this._SN._recordChange({
                type: 'col',
                action: 'outlineLevel',
                sheet: this.sheet,
                col: this.cIndex,
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
        return this._SN._withOperation('colCollapsed', {
            sheet: this.sheet,
            col: this.cIndex,
            oldCollapsed: old,
            newCollapsed: next
        }, () => {
            this._SN.UndoRedo.add({
                undo: () => { this._collapsed = old; },
                redo: () => { this._collapsed = next; }
            });
            this._collapsed = next;
            this._SN._recordChange({
                type: 'col',
                action: 'collapsed',
                sheet: this.sheet,
                col: this.cIndex,
                oldCollapsed: old,
                newCollapsed: next
            });
        });
    }

    /**
     * 是否需要写入 XML
     * @type {boolean}     */
    get _needBuild() {
        if (this.#width !== this.sheet.defaultColWidth || this.#hidden || this.#style || this.#outlineLevel > 0 || this.#collapsed) return true
        return false;
    }

    // ========== 列样式 ==========
    get _colStyle() { return this.#style; }
    set _colStyle(v) {
        this.#style = v;
        if (v) this.sheet._styledCols.add(this.cIndex);
        else this.sheet._styledCols.delete(this.cIndex);
    }

    /**
     * 列样式
     * @type {Object}
     */
    get style() { return this.#style || {}; }
    set style(obj) {
        const old = this.#style;
        const oldInSet = this.sheet._styledCols.has(this.cIndex);
        this._colStyle = obj ? { ...obj } : null;
        this._SN.UndoRedo.add({
            undo: () => {
                this.#style = old;
                if (oldInSet) this.sheet._styledCols.add(this.cIndex);
                else this.sheet._styledCols.delete(this.cIndex);
            },
            redo: () => {
                this._colStyle = obj ? { ...obj } : null;
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
