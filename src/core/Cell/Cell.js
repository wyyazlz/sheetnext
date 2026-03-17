import { CELL_TYPES } from './constants.js';
import { compareValue, dateTrans, numFmtFun, dateStrToDate, formatValue } from './helpers.js';
import { buildCellXml } from './xmlBuilder.js';

/**
 * Cell classes
 * @title Cell Actions
 * @class
 */
export default class Cell {
    /**
     * @param {Object} xmlObj - Cell XML Object
     * @param {Row} row - Row Objects
     * @param {number} cIndex - Column Index
     */
    constructor(xmlObj, row, cIndex) {
        /**
         * Column Index
         * @type {number}
         */
        this.cIndex = cIndex;
        /**
         * Rows belonging to
         * @type {Row}
         */
        this.row = row;
        /**
         * SheetNext Main Instance
         * @type {Object}
         */
        this._SN = row._SN;
        /**
         * Cell XML Object
         * @type {Object}
         */
        this._xmlObj = xmlObj ?? {};
        /**
         * Whether or not to merge cells
         * @type {boolean}
         */
        this.isMerged = false;
        /**
         * Merge Cell Main Cell References
         * @type {{r:number, c:number}|null}
         */
        this.master = null;
        this._hyperlink = null;
        this._dataValidation = null;
        this._style = undefined;
        this._type = undefined;
        this._editVal = undefined;
        this._calcVal = undefined;
        this._showVal = undefined;
        this._fmtResult = undefined; // 完整格式化结果（含会计格式数据）
        this._richText = null;       // 富文本 runs 数组

        // Spill 相关属性
        this._spillRange = null;      // 溢出范围 { rows, cols }
        this._spillParent = null;     // 溢出源单元格引用
        this._spillOffset = null;     // 相对于源单元格的偏移 { r, c }

        // Volatile 函数标记
        this._isVolatile = false;     // 是否包含 volatile 函数（RAND, TODAY 等）

        this._init();
    }

    _init() {
        let { _$t, v, f, is } = this._xmlObj;

        // 初始化类型
        this._type = CELL_TYPES[_$t];
        if (!this._type) {
            if (/[dmyhs]/i.test((this.numFmt ?? "").replaceAll('[Red]', ''))) {
                this._type = undefined;
            } else {
                this._type = 'number';
            }
        }
        const type = this.type;

        // 初始化公式栏值
        if (f) {
            if (typeof f === 'string') {
                this._editVal = `=${f}`;
            } else if (f['_$t'] === 'shared') {
                if (f['#text']) {
                    this._SN._sharedFormula[f['_$si']] = { m: { r: this.row.rIndex, c: this.cIndex }, f: f['#text'] };
                    this._editVal = `=${f['#text']}`;
                } else {
                    const s = this._SN._sharedFormula[f['_$si']];
                    const c = this._SN.Utils.getFillFormula(s.f, s.m.r, s.m.c, this.row.rIndex, this.cIndex);
                    this._editVal = `=${c ?? '公式加载错误'}`;
                }
            } else {
                this._editVal = `=${f?.['#text'] ?? ""}`;
            }
            this._calcVal = ['date', 'dateTime', 'time', 'number'].includes(type) && v ? Number(v) : v;
        } else {
            if (!v && !is) {
                this._editVal = "";
            } else if (_$t === 's') {
                const result = this._SN.Xml._xGetSharedStrings(v);
                this._editVal = result.text;
                if (result.richText) this._richText = result.richText;
            } else if (_$t === 'inlineStr') {
                if (is?.r) {
                    // 内联富文本
                    const rArr = Array.isArray(is.r) ? is.r : [is.r];
                    const runs = [];
                    let hasFormat = false;
                    for (const item of rArr) {
                        const text = typeof item.t === 'string' ? item.t : (item.t?.['#text'] ?? "");
                        const font = this._SN.Xml._parseRunFont(item.rPr);
                        if (font) hasFormat = true;
                        runs.push({ text, font: font || {} });
                    }
                    this._editVal = runs.map(r => r.text).join("");
                    if (hasFormat) this._richText = runs;
                } else {
                    this._editVal = is?.t ?? is;
                }
            } else if (!_$t || _$t == 'n') {
                this._editVal = Number(v);
            } else {
                this._editVal = v;
            }

            if (['date', 'dateTime', 'time'].includes(type) && this._editVal) {
                this._calcVal = this._editVal;
                this._editVal = dateTrans(Number(this._editVal));
            } else {
                this._calcVal = this._editVal;
            }
        }

        if (typeof this._calcVal === 'string') this._calcVal = this._SN.Utils.decodeEntities(this._calcVal);
        if (typeof this._editVal === 'string') this._editVal = this._SN.Utils.decodeEntities(this._editVal);
    }

    set editVal(val) {
        const oldValue = this._editVal;
        const oldType = this._type;

        // 触发 beforeCellEdit 事件（可取消）
        const beforeEvent = this._SN.Event.emit('beforeCellEdit', {
            cell: this,
            sheet: this.row.sheet,
            row: this.row.rIndex,
            col: this.cIndex,
            oldValue,
            newValue: val
        });
        if (beforeEvent.canceled) return;

        // 手动计算模式下保留缓存值（除非是新公式）
        if (this._SN.calcMode !== 'manual') {
            this._calcVal = undefined;
            this._showVal = undefined;
            this._fmtResult = undefined;
        }

        // 编辑纯文本时清除富文本
        if (this._richText) this._richText = null;

        if (val == null || val === '' || val == undefined || Number.isNaN(val)) {
            this._type = undefined;
            this._editVal = "";
        } else if (!isNaN(val)) {
            this._editVal = parseFloat(val);
            this._type = undefined;
        } else if (val?.trim?.().startsWith('=')) {
            this._editVal = val.trim();
            this._type = undefined;
        } else if (/^(true|false)$/i.test(val)) {
            this._editVal = val.toUpperCase() == 'TRUE' ? true : false;
            this._type = 'boolean';
        } else {
            if (typeof val != 'string') debugger;
            const { date, numfmt, type } = (dateStrToDate(val) ?? {});
            if (date) {
                this._editVal = date;
                this._type = type;
                if (!this.style.numFmt) this.style.numFmt = numfmt;
            } else if (/^\d{1,2}:\d{1,2}(:\d{1,2})?$/.test(val)) {
                this._editVal = new Date(`1899/12/31 ${val}`);
                this._type = 'time';
                if (!this.style.numFmt) this.style.numFmt = 'h:mm';
            } else {
                this._editVal = val;
                this._type = 'string';
            }
        }

        const newValue = this._editVal;
        const newType = this._type;

        if (oldValue === newValue && oldType === newType) return;

        const sheet = this.row.sheet;
        this._SN._withOperation('cellEdit', {
            sheet,
            cell: this,
            row: this.row.rIndex,
            col: this.cIndex,
            oldValue,
            newValue,
            oldType,
            newType
        }, () => {

        const wasFormula = typeof oldValue === 'string' && oldValue.startsWith('=');
        const isFormulaNow = this.isFormula;

        // 如果之前是公式，清除 volatile 状态（新公式计算时会重新注册）
        if (wasFormula) {
            this._clearVolatileStatus();
        }

        if (this._SN.DependencyGraph) {
            if (wasFormula || isFormulaNow) this._SN.DependencyGraph.updateDependencies(this);
            if (oldValue !== newValue) {
                const sheetName = this.row.sheet.name;
                this._SN.DependencyGraph.notifyChange(sheetName, this.row.rIndex, this.cIndex);
            }
        }

        if (oldValue !== newValue && this.row.sheet.Sparkline) {
            this.row.sheet.Sparkline.clearAllCache();
        }

        // 清除条件格式缓存
        if (oldValue !== newValue && this.row.sheet.CF) {
            this.row.sheet.CF.clearAllCache();
        }

        // 清除含引用的图表缓存
        if (oldValue !== newValue && this.row.sheet.Drawing) {
            this.row.sheet.Drawing._clearChartRefCaches();
        }

        const undoAction = {
            undo: () => {
                this._editVal = oldValue;
                this._type = oldType;
                this._calcVal = undefined;
                this._showVal = undefined;
                this._fmtResult = undefined;
                if (this._SN.DependencyGraph) {
                    this._SN.DependencyGraph.updateDependencies(this);
                    this._SN.DependencyGraph.notifyChange(this.row.sheet.name, this.row.rIndex, this.cIndex);
                }
                if (this.row.sheet.Sparkline) this.row.sheet.Sparkline.clearAllCache();
                if (this.row.sheet.CF) this.row.sheet.CF.clearAllCache();
                if (this.row.sheet.Drawing) this.row.sheet.Drawing._clearChartRefCaches();
            },
            redo: () => {
                this._editVal = newValue;
                this._type = newType;
                this._calcVal = undefined;
                this._showVal = undefined;
                this._fmtResult = undefined;
                if (this._SN.DependencyGraph) {
                    this._SN.DependencyGraph.updateDependencies(this);
                    this._SN.DependencyGraph.notifyChange(this.row.sheet.name, this.row.rIndex, this.cIndex);
                }
                if (this.row.sheet.Sparkline) this.row.sheet.Sparkline.clearAllCache();
                if (this.row.sheet.CF) this.row.sheet.CF.clearAllCache();
                if (this.row.sheet.Drawing) this.row.sheet.Drawing._clearChartRefCaches();
            }
        };
        if (sheet === this._SN.activeSheet) {
            undoAction.activeCell = { r: this.row.rIndex, c: this.cIndex };
        }
        this._SN.UndoRedo.add(undoAction);

        this._SN._recordChange({
            type: 'cellEdit',
            sheet,
            row: this.row.rIndex,
            col: this.cIndex,
            oldValue,
            newValue,
            oldType,
            newType
        });

        // 触发 afterCellEdit 事件
        this._SN.Event.emit('afterCellEdit', {
            cell: this,
            sheet,
            row: this.row.rIndex,
            col: this.cIndex,
            oldValue,
            newValue
        });
        }, { cancelable: false });
    }

    /**
     * Edit value or formula
     * @type {string}
     */
    get editVal() {
        if (this._editVal instanceof Date && this.numFmt) {
            const type = this.type;
            if (type == 'date') return this._editVal.toLocaleDateString();
            else if (type == 'time') return this._editVal.toLocaleTimeString();
            else return this._editVal.toLocaleString();
        }
        return this._editVal?.toString() ?? "";
    }

    /**
     * Rich text runs array
     * @type {Array|null}
     */
    get richText() { return this._richText; }
    set richText(runs) {
        const old = this._richText;
        const oldEditVal = this._editVal;
        this._richText = runs;
        this._showVal = undefined;
        this._calcVal = undefined;
        this._fmtResult = undefined;
        // 同步 _editVal 为所有 run 的文本拼接，确保导出/搜索/公式引用正常
        if (runs) {
            this._editVal = runs.map(r => r.text).join('');
        }
        this._SN.UndoRedo.add({
            undo: () => { this._richText = old; this._editVal = oldEditVal; this._showVal = undefined; this._calcVal = undefined; this._fmtResult = undefined; },
            redo: () => { this._richText = runs; this._editVal = runs ? runs.map(r => r.text).join('') : oldEditVal; this._showVal = undefined; this._calcVal = undefined; this._fmtResult = undefined; }
        });
    }

    /**
     * Show values
     * @type {string}     */
    /**
     * Show values
     * @type {string}     */

    get showVal() {
        if (this._showVal != undefined) return this._showVal;
        if (this.type == 'boolean') return this._showVal = this.calcVal ? 'TRUE' : 'FALSE';
        else if (this._type == 'error') return this._showVal = this.calcVal ?? "";
        else if (this.numFmt || ['date', 'time', 'dateTime'].includes(this._type)) {
            // 使用 formatValue 获取完整结果（含会计格式数据）
            this._fmtResult = formatValue(this.calcVal, this.numFmt, this._type);
            return this._showVal = (this._fmtResult?.text ?? "") + "";
        }
        else return this._showVal = String(this.calcVal);
    }

    /**
     * Get Accounting Format Structured Data (for Rendering)
     * @type {Object|undefined}     */
    /**
     * Get Accounting Format Structured Data (for Rendering)
     * @type {Object|undefined}     */

    get accountingData() {
        if (this._showVal === undefined) this.showVal; // 触发计算
        return this._fmtResult?.accounting;
    }

    /**
     * Calculated value
     * @type {any}     */
    /**
     * Calculated value
     * @type {any}     */

    get calcVal() {
        if (this._calcVal != undefined) return this._calcVal;

        // 如果是 spill 引用单元格，返回源单元格的对应值
        if (this._spillParent && this._spillOffset) {
            const parentVal = this._spillParent.calcVal;
            if (Array.isArray(parentVal) && Array.isArray(parentVal[0])) {
                const { r, c } = this._spillOffset;
                return this._calcVal = parentVal[r]?.[c] ?? '';
            }
            return this._calcVal = '';
        }

        if (this.isFormula) {
            let calcRes = this._SN.Formula.calcFormula(this.editVal.trim().substring(1), this);

            // 检测并注册 Volatile 状态
            this._updateVolatileStatus();

            // 处理数组结果（Spill）
            if (Array.isArray(calcRes) && Array.isArray(calcRes[0])) {
                const rows = calcRes.length;
                const cols = calcRes[0].length;

                // 检查是否需要 spill（超过 1x1）
                if (rows > 1 || cols > 1) {
                    // 检查溢出区域是否有冲突
                    const sheet = this.row.sheet;
                    let hasConflict = false;

                    for (let r = 0; r < rows && !hasConflict; r++) {
                        for (let c = 0; c < cols && !hasConflict; c++) {
                            if (r === 0 && c === 0) continue; // 跳过源单元格
                            const targetCell = sheet.getCell(this.row.rIndex + r, this.cIndex + c);
                            // 检查目标单元格是否有数据（非空且不是当前公式的 spill）
                            if (targetCell && targetCell._editVal !== undefined && targetCell._editVal !== '' &&
                                targetCell._spillParent !== this) {
                                hasConflict = true;
                            }
                        }
                    }

                    if (hasConflict) {
                        return this._calcVal = '#SPILL!';
                    }

                    // 设置 spill 范围
                    this._spillRange = { rows, cols };

                    // 清除旧的 spill 引用
                    this._clearSpillRefs();

                    // 设置 spill 引用单元格
                    for (let r = 0; r < rows; r++) {
                        for (let c = 0; c < cols; c++) {
                            if (r === 0 && c === 0) continue;
                            const targetCell = sheet.getCell(this.row.rIndex + r, this.cIndex + c);
                            if (targetCell) {
                                targetCell._spillParent = this;
                                targetCell._spillOffset = { r, c };
                                targetCell._calcVal = undefined; // 重置计算值
                                targetCell._showVal = undefined;
                            }
                        }
                    }

                    // 返回左上角的值
                    return this._calcVal = calcRes[0][0];
                }

                // 1x1 数组，直接返回值
                return this._calcVal = calcRes[0][0];
            }

            // 检查公式结果是否为日期类型（通过标记判断，性能更优）
            if (this._SN.Formula.lastResultIsDate && typeof calcRes === 'number') {
                if (!this.numFmt) {
                    if (calcRes < 1) { this.numFmt = 'h:mm'; this._type = 'time'; }
                    else if (Number.isInteger(calcRes)) { this.numFmt = 'yyyy/m/d'; this._type = 'date'; }
                    else { this.numFmt = 'yyyy/m/d h:mm:ss'; this._type = 'dateTime'; }
                }
            }
            if (calcRes instanceof Error) {
                this._type = 'error';
                return this._calcVal = calcRes.message;
            }
            return this._calcVal = calcRes;
        } else {
            if (this._editVal instanceof Date) return this._calcVal = dateTrans(this._editVal);
            else return this._calcVal = this._editVal;
        }
    }

    // 清除当前单元格的 spill 引用
    _clearSpillRefs() {
        if (!this._spillRange) return;
        const sheet = this.row.sheet;
        const { rows, cols } = this._spillRange;
        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                if (r === 0 && c === 0) continue;
                const targetCell = sheet.getCell(this.row.rIndex + r, this.cIndex + c);
                if (targetCell && targetCell._spillParent === this) {
                    targetCell._spillParent = null;
                    targetCell._spillOffset = null;
                    targetCell._calcVal = undefined;
                    targetCell._showVal = undefined;
                }
            }
        }
        this._spillRange = null;
    }

    // 更新 Volatile 状态（根据公式计算结果注册/注销）
    _updateVolatileStatus() {
        if (!this._SN.DependencyGraph) return;

        const isVolatile = this._SN.Formula.lastCalcVolatile;
        const sheetName = this.row.sheet.name;

        if (isVolatile && !this._isVolatile) {
            // 新增 volatile 状态
            this._isVolatile = true;
            this._SN.DependencyGraph.registerVolatile(sheetName, this.row.rIndex, this.cIndex);
        } else if (!isVolatile && this._isVolatile) {
            // 移除 volatile 状态
            this._isVolatile = false;
            this._SN.DependencyGraph.unregisterVolatile(sheetName, this.row.rIndex, this.cIndex);
        }
    }

    // 清除 Volatile 状态（在公式被删除或修改时调用）
    _clearVolatileStatus() {
        if (this._isVolatile && this._SN.DependencyGraph) {
            this._isVolatile = false;
            this._SN.DependencyGraph.unregisterVolatile(
                this.row.sheet.name,
                this.row.rIndex,
                this.cIndex
            );
        }
    }

    /**
     * Get full spill array results
     * @type {Array[]|null}     */
    /**
     * Get full spill array results
     * @type {Array[]|null}     */

    get spillArray() {
        if (!this._spillRange || !this.isFormula) return null;
        const calcRes = this._SN.Formula.calcFormula(this.editVal.trim().substring(1), this);
        return Array.isArray(calcRes) ? calcRes : null;
    }

    /**
     * Whether it is a spill source cell
     * @type {boolean}     */
    /**
     * Whether it is a spill source cell
     * @type {boolean}     */

    get isSpillSource() {
        return this._spillRange !== null;
    }

    /**
     * Whether it is a spill reference cell
     * @type {boolean}     */
    /**
     * Whether it is a spill reference cell
     * @type {boolean}     */

    get isSpillRef() {
        return this._spillParent !== null;
    }

    /**
     * Calculated Horizontal Alignment
     * @type {string}     */
    /**
     * Calculated Horizontal Alignment
     * @type {string}     */

    get horizontalAlign() {
        if (this.alignment?.horizontal) return this.alignment.horizontal;
        if (this.numFmt === '@') return 'left';
        if (this._editVal === '' || this._editVal == null) return 'left';
        const t = this.type;
        if (['date', 'dateTime', 'time', 'number'].includes(t)) return 'right';
        else if (['error', 'boolean'].includes(t)) return 'center';
        else return 'left';
    }

    /**
     * Calculated Vertical Alignment
     * @type {string}     */
    /**
     * Calculated Vertical Alignment
     * @type {string}     */

    get verticalAlign() {
        return this.alignment?.vertical || 'center';
    }

    /**
     * Hyperlink Configuration
     * @type {Object|null}
     */
    get hyperlink() { return this._hyperlink; }
    set hyperlink(obj) {
        const previousState = this._hyperlink ? { ...this._hyperlink } : null;
        const isInsert = !previousState && obj && Object.keys(obj).length > 0;

        // 插入超链接时触发事件
        if (isInsert && this._SN.Event.hasListeners('beforeHyperlinkInsert')) {
            const beforeEvent = this._SN.Event.emit('beforeHyperlinkInsert', {
                sheet: this.row.sheet,
                cell: this,
                row: this.row.rIndex,
                col: this.cIndex,
                hyperlink: obj
            });
            if (beforeEvent.canceled) return;
        }

        if (!obj || Object.keys(obj).length == 0) {
            this._hyperlink = null;
        } else {
            this._hyperlink = { ...obj };
            this.font = { color: "#336699", underline: 'single' };
        }
        const currentState = this._hyperlink ? { ...this._hyperlink } : null;
        this._SN.UndoRedo.add({
            undo: () => { this._hyperlink = previousState; },
            redo: () => { this._hyperlink = currentState; }
        });

        if (isInsert && this._SN.Event.hasListeners('afterHyperlinkInsert')) {
            this._SN.Event.emit('afterHyperlinkInsert', {
                sheet: this.row.sheet,
                cell: this,
                row: this.row.rIndex,
                col: this.cIndex,
                hyperlink: currentState
            });
        }
    }

    /**
     * Data Validation Configuration
     * @type {Object|null}
     */
    get dataValidation() { return this._dataValidation; }
    set dataValidation(obj) {
        const p = JSON.parse(JSON.stringify(this._dataValidation));
        let newVal = null;
        if (obj && Object.keys(obj).length > 0) {
            newVal = obj.type == 'list'
                ? { showDropDown: true, showErrorMessage: true, showInputMessage: true, ...obj }
                : { showInputMessage: true, showErrorMessage: true, ...obj };
        }
        this._dataValidation = newVal;
        this._SN.UndoRedo.add({
            undo: () => { this._dataValidation = p; },
            redo: () => { this._dataValidation = newVal; }
        });
    }

    /**
     * Cell Protection Configuration
     * @type {{locked: boolean, hidden: boolean}}
     */
    get protection() {
        const protection = this.style?.protection;
        if (!protection) return { locked: true, hidden: false };
        return {
            locked: protection.locked !== false,
            hidden: !!protection.hidden
        };
    }
    set protection(obj) {
        const oldProtection = this._style?.protection ? { ...this._style.protection } : null;
        let nextProtection = null;

        if (obj && typeof obj === 'object') {
            const locked = Object.prototype.hasOwnProperty.call(obj, 'locked') ? !!obj.locked : undefined;
            const hidden = Object.prototype.hasOwnProperty.call(obj, 'hidden') ? !!obj.hidden : undefined;
            if (locked === false || hidden === true) {
                nextProtection = {};
                if (locked !== undefined) nextProtection.locked = locked;
                if (hidden !== undefined) nextProtection.hidden = hidden;
            }
        }

        if (!this._style) this._style = {};
        if (nextProtection) {
            this._style.protection = nextProtection;
        } else if (this._style?.protection) {
            delete this._style.protection;
        }

        const currentProtection = this._style?.protection ? { ...this._style.protection } : null;
        this._SN.UndoRedo.add({
            undo: () => {
                if (!this._style) this._style = {};
                if (oldProtection) this._style.protection = { ...oldProtection };
                else delete this._style.protection;
            },
            redo: () => {
                if (!this._style) this._style = {};
                if (currentProtection) this._style.protection = { ...currentProtection };
                else delete this._style.protection;
            }
        });
    }

    /**
     * Is it locked
     * @type {boolean}     */
    /**
     * Is it locked
     * @type {boolean}     */

    get isLocked() {
        return this.protection.locked !== false;
    }

    /**
     * Data validation results
     * @type {boolean}     */
    /**
     * Data validation results
     * @type {boolean}     */

    get validData() {
        if (!this._dataValidation) return true;
        const val = this.calcVal;
        const { type, operator, formula1, formula2, allowBlank } = this._dataValidation;
        const needsTwoBounds = ['whole', 'decimal', 'date', 'time', 'textLength'];
        const op = operator || (needsTwoBounds.includes(type) && formula2 ? 'between' : operator);

        if (allowBlank && (val === '' || val === null || val === undefined)) return true;

        switch (type) {
            case 'list': return formula1.includes(val.toString());
            case 'whole':
                if (!Number.isInteger(Number(val))) return false;
                return compareValue(Number(val), op, formula1, formula2);
            case 'decimal':
                if (isNaN(Number(val))) return false;
                return compareValue(Number(val), op, formula1, formula2);
            case 'textLength': return compareValue(val.toString().length, op, formula1, formula2);
            case 'custom': return true;
            case 'date':
                if (!(val instanceof Date)) return false;
                return compareValue(val.getTime(), op, new Date(formula1).getTime(), new Date(formula2).getTime());
            case 'time':
                if (!(val instanceof Date)) return false;
                const secondsVal = val.getHours() * 3600 + val.getMinutes() * 60 + val.getSeconds();
                const f1 = new Date('1970/1/1 ' + formula1);
                const f2 = new Date('1970/1/1 ' + formula2);
                return compareValue(secondsVal, op, f1.getHours() * 3600 + f1.getMinutes() * 60 + f1.getSeconds(), f2.getHours() * 3600 + f2.getMinutes() * 60 + f2.getSeconds());
            default: return true;
        }
    }

    /**
     * Is it a formula
     * @type {boolean}     */
    /**
     * Is it a formula
     * @type {boolean}     */

    get isFormula() { return typeof this._editVal == 'string' && this._editVal.startsWith("="); }

    /**
     * Cell Type
     * @type {string}
     */
    get type() {
        if (this._type) return this._type;
        const calcVal = this.calcVal;
        if (typeof calcVal === 'boolean') return this._type = 'boolean';
        else if (/[dmyhs]/i.test(this.numFmt ?? "")) {
            const n = this.numFmt;
            const t = n.includes('h') || n.includes('s');
            const d = n.includes('y') || n.includes('d');
            if (t && d) return this._type = 'dateTime';
            else if (t) return this._type = 'time';
            else if (d) return this._type = 'date';
            return this._type = 'dateTime';
        } else if (calcVal != null && !(typeof calcVal === 'string' && calcVal.trim() === '') && !isNaN(calcVal)) return this._type = 'number';
        else if (this._editVal instanceof Error) return this._type = 'error';
        else return this._type = 'string';
    }
    set type(val) { this._type = val; }

    /**
     * Cell Styles
     * @type {Object}
     */
    get style() {
        if (this._style != undefined) return this._style;
        if (this._xmlObj['_$s']) return this._style = this._SN.Xml._xGetStyle(this._xmlObj['_$s']);
        return this._style = {};
    }
    set style(obj) {
        const oldStyle = JSON.parse(JSON.stringify(this._style ?? {}));
        // 触发样式变更事件
        if (this._SN.Event.hasListeners('beforeCellStyleChange')) {
            const beforeEvent = this._SN.Event.emit('beforeCellStyleChange', {
                sheet: this.row.sheet,
                cell: this,
                row: this.row.rIndex,
                col: this.cIndex,
                oldStyle,
                newStyle: obj
            });
            if (beforeEvent.canceled) return;
        }
        this._style = obj ?? {};
        this._SN.UndoRedo.add({
            undo: () => { this._style = oldStyle; },
            redo: () => { this._style = obj; }
        });
        if (this._SN.Event.hasListeners('afterCellStyleChange')) {
            this._SN.Event.emit('afterCellStyleChange', {
                sheet: this.row.sheet,
                cell: this,
                row: this.row.rIndex,
                col: this.cIndex,
                oldStyle,
                newStyle: obj
            });
        }
    }

    /**
     * Border Style
     * @type {Object}
     */
    get border() {
        const c = this.style?.border;
        if (c) return c;
        const sheet = this.row.sheet;
        if (!sheet._styledRows.size && !sheet._styledCols.size) return {};
        const r = sheet._styledRows.has(this.row.rIndex) ? this.row._rowStyle?.border : null;
        const col = sheet._styledCols.has(this.cIndex) ? sheet.cols[this.cIndex]?._colStyle?.border : null;
        if (!r && !col) return {};
        if (r && !col) return r;
        if (!r && col) return col;
        return { ...col, ...r };
    }
    set border(obj) {
        if (Object.keys(obj ?? {}).length == 0) obj = null;
        const oldBorder = JSON.parse(JSON.stringify(this.style.border ?? {}));
        if (!obj) {
            delete this.style.border;
        } else {
            const b = this.style.border = this.style.border || {};
            Object.keys(obj).forEach((name) => {
                const value = obj[name];
                if (['top', 'right', 'bottom', 'left', 'diagonal'].includes(name)) {
                    if (value && typeof value === 'object') {
                        b[name] = { color: '#000000', style: 'thin', ...(b[name] ?? {}), ...value };
                    } else {
                        delete b[name];
                    }
                    return;
                }
                if (name === 'diagonalUp' || name === 'diagonalDown') {
                    if (value === undefined || value === null) delete b[name];
                    else b[name] = !!value;
                    return;
                }
                if (value === undefined || value === null) delete b[name];
                else b[name] = value;
            });
            if (!Object.keys(b).length) delete this.style.border;
        }
        const newBorder = JSON.parse(JSON.stringify(this.style.border ?? {}));
        this._SN.UndoRedo.add({
            undo: () => { this.style.border = oldBorder; },
            redo: () => { this.style.border = newBorder; }
        });
    }

    /**
     * Fill Style
     * @type {Object}
     */
    get fill() {
        const c = this.style?.fill;
        if (c) return c;
        const sheet = this.row.sheet;
        if (!sheet._styledRows.size && !sheet._styledCols.size) return {};
        const r = sheet._styledRows.has(this.row.rIndex) ? this.row._rowStyle?.fill : null;
        const col = sheet._styledCols.has(this.cIndex) ? sheet.cols[this.cIndex]?._colStyle?.fill : null;
        if (!r && !col) return {};
        if (r && !col) return r;
        if (!r && col) return col;
        return { ...col, ...r };
    }
    set fill(obj) {
        if (Object.keys(obj ?? {}).length == 0) obj = null;
        const oldFill = JSON.parse(JSON.stringify(this.style.fill ?? {}));
        if (!obj) delete this._style.fill;
        else {
            const nextFill = { ...(this._style.fill ?? {}), ...obj };
            const hasFgColor = Object.prototype.hasOwnProperty.call(obj, 'fgColor') &&
                obj.fgColor !== undefined &&
                obj.fgColor !== null &&
                obj.fgColor !== '';
            const hasType = Object.prototype.hasOwnProperty.call(obj, 'type');
            const hasPattern = Object.prototype.hasOwnProperty.call(obj, 'pattern');

            // Setting a foreground color should default to a visible solid pattern.
            if (hasFgColor) {
                if (!hasType) nextFill.type = 'pattern';
                if (!hasPattern) nextFill.pattern = 'solid';
            }
            if (!nextFill.type && (nextFill.pattern || nextFill.fgColor || nextFill.bgColor)) {
                nextFill.type = 'pattern';
            }
            if (nextFill.type === 'pattern' && !nextFill.pattern) {
                nextFill.pattern = 'solid';
            }
            this._style.fill = nextFill;
        }
        const newFill = JSON.parse(JSON.stringify(this._style.fill ?? {}));
        this._SN.UndoRedo.add({
            undo: () => { this.style.fill = oldFill; },
            redo: () => { this.style.fill = newFill; }
        });
    }

    /**
     * Font style
     * @type {Object}
     */
    get font() {
        const c = this.style?.font;
        if (c) return c;
        const sheet = this.row.sheet;
        if (!sheet._styledRows.size && !sheet._styledCols.size) return {};
        const r = sheet._styledRows.has(this.row.rIndex) ? this.row._rowStyle?.font : null;
        const col = sheet._styledCols.has(this.cIndex) ? sheet.cols[this.cIndex]?._colStyle?.font : null;
        if (!r && !col) return {};
        if (r && !col) return r;
        if (!r && col) return col;
        return { ...col, ...r };
    }
    set font(obj) {
        if (Object.keys(obj ?? {}).length == 0) obj = null;
        const oldFont = JSON.parse(JSON.stringify(this.style.font ?? {}));
        if (!obj) delete this.style.font;
        else {
            const nextFont = { ...this._style.font, ...obj };
            if (Object.prototype.hasOwnProperty.call(obj, 'color')) {
                delete nextFont.colorTheme;
                delete nextFont.colorTint;
                delete nextFont.colorIndexed;
                delete nextFont.colorAuto;
            }
            this.style.font = nextFont;
        }
        const newFont = JSON.parse(JSON.stringify(this.style.font ?? {}));
        this._SN.UndoRedo.add({
            undo: () => { this.style.font = oldFont; },
            redo: () => { this.style.font = newFont; }
        });
    }

    /**
     * Alignment
     * @type {Object}
     */
    get alignment() {
        const c = this.style?.alignment;
        if (c) return c;
        const sheet = this.row.sheet;
        if (!sheet._styledRows.size && !sheet._styledCols.size) return {};
        const r = sheet._styledRows.has(this.row.rIndex) ? this.row._rowStyle?.alignment : null;
        const col = sheet._styledCols.has(this.cIndex) ? sheet.cols[this.cIndex]?._colStyle?.alignment : null;
        if (!r && !col) return {};
        if (r && !col) return r;
        if (!r && col) return col;
        return { ...col, ...r };
    }
    set alignment(obj) {
        if (Object.keys(obj ?? {}).length == 0) obj = null;
        const oldAlignment = JSON.parse(JSON.stringify(this.style.alignment ?? {}));
        if (!obj) delete this.style.alignment;
        else this.style.alignment = { ...this._style.alignment, ...obj };
        const newAlignment = JSON.parse(JSON.stringify(this.style.alignment ?? {}));
        this._SN.UndoRedo.add({
            undo: () => { this.style.alignment = oldAlignment; },
            redo: () => { this.style.alignment = newAlignment; }
        });
    }

    /**
     * Number Format
     * @type {string|undefined}
     */
    get numFmt() {
        const c = this.style?.numFmt;
        if (c) return c;
        const sheet = this.row.sheet;
        if (!sheet._styledRows.size && !sheet._styledCols.size) return undefined;
        const r = sheet._styledRows.has(this.row.rIndex) ? this.row._rowStyle?.numFmt : null;
        if (r) return r;
        const col = sheet._styledCols.has(this.cIndex) ? sheet.cols[this.cIndex]?._colStyle?.numFmt : null;
        return col || undefined;
    }
    set numFmt(fmtStr) {
        const oldNumFmt = this.style.numFmt;
        const oldShowVal = this._showVal;
        const oldFmtResult = this._fmtResult;
        if (!fmtStr) delete this.style.numFmt;
        else this.style.numFmt = fmtStr;
        this._showVal = undefined;
        this._fmtResult = undefined;
        const newNumFmt = this.style.numFmt;
        const newShowVal = this._showVal;
        const newFmtResult = this._fmtResult;
        this._SN.UndoRedo.add({
            undo: () => { this.style.numFmt = oldNumFmt; this._showVal = oldShowVal; this._fmtResult = oldFmtResult; },
            redo: () => { this.style.numFmt = newNumFmt; this._showVal = newShowVal; this._fmtResult = newFmtResult; }
        });
    }

    /**
     * Build Cell XML
     * @type {Object}     */
    /**
     * Build Cell XML
     * @type {Object}     */

    get buildXml() { return buildCellXml(this); }

    /**
     * Do you need to build XML
     * @type {boolean}     */
    /**
     * Do you need to build XML
     * @type {boolean}     */

    get _needBuild() {
        return ![undefined, null, ""].includes(this.editVal) || Object.keys(this.style).length > 0 || this._dataValidation || this._hyperlink;
    }
}
