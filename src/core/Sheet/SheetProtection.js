// 工作表保护类 - 提供优雅的 get/set API 和事件监听控制

import { EventPriority } from '../Event/EventEmitter.js';

const PROTECTION_KEYS = [
    'selectLockedCells',
    'selectUnlockedCells',
    'formatCells',
    'formatColumns',
    'formatRows',
    'insertColumns',
    'insertRows',
    'insertHyperlinks',
    'deleteColumns',
    'deleteRows',
    'sort',
    'autoFilter',
    'pivotTables',
    'objects'
];

const PROTECTION_DEFAULTS = {
    selectLockedCells: true,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false,
    insertColumns: false,
    insertRows: false,
    insertHyperlinks: false,
    deleteColumns: false,
    deleteRows: false,
    sort: false,
    autoFilter: false,
    pivotTables: false,
    objects: false
};

// 密码哈希（Excel 兼容的弱哈希算法）
function hashPassword(password) {
    const text = String(password ?? '');
    if (!text) return '';

    let hash = 0;
    for (let i = 0; i < text.length; i++) {
        const charCode = text.charCodeAt(i);
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= charCode;
    }
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
    hash ^= text.length;
    hash ^= 0xCE4B;

    return hash.toString(16).toUpperCase().padStart(4, '0');
}

// 解析 xlsx 中的 sheetProtection 节点
export function parseSheetProtectionNode(node) {
    if (!node) return null;
    const sheetEnabled = toBoolean(node._$sheet, true);
    if (!sheetEnabled) return null;

    const options = { sheet: true };
    PROTECTION_KEYS.forEach((key) => {
        options[key] = toBoolean(node[`_$${key}`], PROTECTION_DEFAULTS[key]);
    });

    if (node._$password) options.passwordHash = String(node._$password).toUpperCase();
    if (node._$algorithmName) options.algorithmName = node._$algorithmName;
    if (node._$hashValue) options.hashValue = node._$hashValue;
    if (node._$saltValue) options.saltValue = node._$saltValue;
    if (node._$spinCount !== undefined) options.spinCount = Number(node._$spinCount);

    return options;
}

function toBoolean(value, fallback) {
    if (value === undefined || value === null) return fallback;
    if (value === true || value === false) return value;
    if (value === 1 || value === '1') return true;
    if (value === 0 || value === '0') return false;
    return fallback;
}

export default class SheetProtection {
    constructor(sheet) {
        this._sheet = sheet;
        this._enabled = false;
        this._options = { ...PROTECTION_DEFAULTS };
        this._passwordHash = null;
        this._algorithmName = null;
        this._hashValue = null;
        this._saltValue = null;
        this._spinCount = null;
        this._listenersBound = false;
    }

    get SN() {
        return this._sheet.SN;
    }

    _t(key, params) {
        return this.SN.t(key, params);
    }

    // ==================== 核心属性 ====================

    get enabled() {
        return this._enabled;
    }

    // ==================== 15个标准点位 get/set ====================

    get selectLockedCells() { return this._options.selectLockedCells; }
    set selectLockedCells(v) { this._setOption('selectLockedCells', !!v); }

    get selectUnlockedCells() { return this._options.selectUnlockedCells; }
    set selectUnlockedCells(v) { this._setOption('selectUnlockedCells', !!v); }

    get formatCells() { return this._options.formatCells; }
    set formatCells(v) { this._setOption('formatCells', !!v); }

    get formatColumns() { return this._options.formatColumns; }
    set formatColumns(v) { this._setOption('formatColumns', !!v); }

    get formatRows() { return this._options.formatRows; }
    set formatRows(v) { this._setOption('formatRows', !!v); }

    get insertColumns() { return this._options.insertColumns; }
    set insertColumns(v) { this._setOption('insertColumns', !!v); }

    get insertRows() { return this._options.insertRows; }
    set insertRows(v) { this._setOption('insertRows', !!v); }

    get insertHyperlinks() { return this._options.insertHyperlinks; }
    set insertHyperlinks(v) { this._setOption('insertHyperlinks', !!v); }

    get deleteColumns() { return this._options.deleteColumns; }
    set deleteColumns(v) { this._setOption('deleteColumns', !!v); }

    get deleteRows() { return this._options.deleteRows; }
    set deleteRows(v) { this._setOption('deleteRows', !!v); }

    get sort() { return this._options.sort; }
    set sort(v) { this._setOption('sort', !!v); }

    get autoFilter() { return this._options.autoFilter; }
    set autoFilter(v) { this._setOption('autoFilter', !!v); }

    get pivotTables() { return this._options.pivotTables; }
    set pivotTables(v) { this._setOption('pivotTables', !!v); }

    get objects() { return this._options.objects; }
    set objects(v) { this._setOption('objects', !!v); }

    // ==================== 核心方法 ====================

    /**
     * 启用工作表保护
     * @param {Object} options - 保护选项
     * @param {string} [options.password] - 明文密码（会自动哈希）
     * @param {string} [options.passwordHash] - 已哈希的密码
     * @param {boolean} [options.selectLockedCells] - 允许选择锁定单元格
     * @param {boolean} [options.selectUnlockedCells] - 允许选择未锁定单元格
     * @param {boolean} [options.formatCells] - 允许设置单元格格式
     * @param {boolean} [options.formatColumns] - 允许设置列格式
     * @param {boolean} [options.formatRows] - 允许设置行格式
     * @param {boolean} [options.insertColumns] - 允许插入列
     * @param {boolean} [options.insertRows] - 允许插入行
     * @param {boolean} [options.insertHyperlinks] - 允许插入超链接
     * @param {boolean} [options.deleteColumns] - 允许删除列
     * @param {boolean} [options.deleteRows] - 允许删除行
     * @param {boolean} [options.sort] - 允许排序
     * @param {boolean} [options.autoFilter] - 允许使用自动筛选
     * @param {boolean} [options.pivotTables] - 允许使用数据透视表
     * @param {boolean} [options.objects] - 允许编辑对象
     */
    enable(options = {}) {
        // 设置选项
        PROTECTION_KEYS.forEach((key) => {
            if (options[key] !== undefined) {
                this._options[key] = !!options[key];
            } else {
                this._options[key] = PROTECTION_DEFAULTS[key];
            }
        });

        // 处理密码
        if (options.passwordHash) {
            this._passwordHash = String(options.passwordHash).toUpperCase();
        } else if (options.password) {
            this._passwordHash = hashPassword(options.password);
        } else {
            this._passwordHash = null;
        }

        // 高级加密选项
        this._algorithmName = options.algorithmName || null;
        this._hashValue = options.hashValue || null;
        this._saltValue = options.saltValue || null;
        this._spinCount = options.spinCount !== undefined ? Number(options.spinCount) : null;

        this._enabled = true;
        this._registerListeners();

        return this;
    }

    /**
     * 禁用工作表保护
     * @param {string} [password] - 密码（如果设置了密码保护）
     * @returns {{ ok: boolean, message?: string }}
     */
    disable(password) {
        if (!this._enabled) return { ok: true };

        // 验证密码
        const verify = this.verifyPassword(password);
        if (!verify.ok) return verify;

        this._unregisterListeners();
        this._enabled = false;
        this._passwordHash = null;
        this._algorithmName = null;
        this._hashValue = null;
        this._saltValue = null;
        this._spinCount = null;

        return { ok: true };
    }

    /**
     * 验证密码
     * @param {string} input - 输入的密码
     * @returns {{ ok: boolean, message?: string }}
     */
    verifyPassword(input) {
        if (!this._enabled) return { ok: true };
        if (!this._passwordHash && !this._hashValue) return { ok: true };

        const raw = String(input ?? '');
        if (!raw) return { ok: false, message: this._t('core.sheet.sheetProtection.prompt.enterPassword') };

        if (this._passwordHash) {
            const hash = hashPassword(raw);
            if (hash === this._passwordHash) return { ok: true };
            return { ok: false, message: this._t('core.sheet.sheetProtection.prompt.passwordIncorrect') };
        }

        return { ok: false, message: this._t('core.sheet.sheetProtection.prompt.advancedNotSupported') };
    }

    /**
     * 获取所有保护选项（用于 UI 显示或导出）
     * @returns {Object|null}
     */
    getOptions() {
        if (!this._enabled) return null;
        return {
            sheet: true,
            ...this._options,
            passwordHash: this._passwordHash,
            algorithmName: this._algorithmName,
            hashValue: this._hashValue,
            saltValue: this._saltValue,
            spinCount: this._spinCount
        };
    }

    /**
     * 检查单元格是否锁定
     * @param {Object} cell - 单元格对象
     * @returns {boolean}
     */
    isCellLocked(cell) {
        if (!cell) return true;
        return cell.protection?.locked !== false;
    }

    /**
     * 检查单元格是否可选择
     * @param {Object} cell - 单元格对象
     * @returns {boolean}
     */
    isCellSelectable(cell) {
        if (!this._enabled) return true;
        const allowLocked = this._options.selectLockedCells !== false;
        const allowUnlocked = this._options.selectUnlockedCells !== false;
        if (!allowLocked && !allowUnlocked) return false;
        const locked = this.isCellLocked(cell);
        return locked ? allowLocked : allowUnlocked;
    }

    /**
     * 检查区域是否可选择
     * @param {Object|string} area - 区域对象或字符串
     * @returns {boolean}
     */
    isAreaSelectable(area) {
        if (!this._enabled) return true;
        const allowLocked = this._options.selectLockedCells !== false;
        const allowUnlocked = this._options.selectUnlockedCells !== false;
        if (!allowLocked && !allowUnlocked) return false;

        const range = typeof area === 'string' ? this._sheet.rangeStrToNum(area) : area;
        if (!range?.s || !range?.e) return true;

        const sr = Math.min(range.s.r, range.e.r);
        const er = Math.max(range.s.r, range.e.r);
        const sc = Math.min(range.s.c, range.e.c);
        const ec = Math.max(range.s.c, range.e.c);

        for (let r = sr; r <= er; r++) {
            for (let c = sc; c <= ec; c++) {
                const cell = this._sheet.getCell(r, c);
                const locked = this.isCellLocked(cell);
                if ((locked && !allowLocked) || (!locked && !allowUnlocked)) {
                    return false;
                }
            }
        }
        return true;
    }

    /**
     * 检查多个区域是否可选择
     * @param {Array} areas - 区域数组
     * @returns {boolean}
     */
    areAreasSelectable(areas) {
        if (!this._enabled) return true;
        if (!areas || areas.length === 0) return true;
        return areas.every(area => this.isAreaSelectable(area));
    }

    /**
     * 是否允许显示选区
     * @returns {boolean}
     */
    isSelectionVisible() {
        if (!this._enabled) return true;
        return this._options.selectLockedCells !== false || this._options.selectUnlockedCells !== false;
    }

    /**
     * 设置单元格保护属性
     * @param {Object|string|Array} area - 区域
     * @param {Object} options - { locked?: boolean, hidden?: boolean }
     */
    setCellProtection(area, options = {}) {
        const regions = area ? (Array.isArray(area) ? area : [area]) : this._sheet.activeAreas;
        const hasLocked = Object.prototype.hasOwnProperty.call(options, 'locked');
        const hasHidden = Object.prototype.hasOwnProperty.call(options, 'hidden');
        if (!hasLocked && !hasHidden) return;

        regions.forEach((region) => {
            const range = typeof region === 'string' ? this._sheet.rangeStrToNum(region) : region;
            this._sheet.eachCells(range, (r, c) => {
                const cell = this._sheet.getCell(r, c);
                const current = cell.protection || {};
                cell.protection = {
                    locked: hasLocked ? !!options.locked : current.locked,
                    hidden: hasHidden ? !!options.hidden : current.hidden
                };
            });
        });
    }

    /**
     * 构建 xlsx 导出用的 sheetProtection 节点
     * @returns {Object|null}
     */
    buildXmlNode() {
        if (!this._enabled) return null;

        const node = { '_$sheet': '1' };
        PROTECTION_KEYS.forEach((key) => {
            node[`_$${key}`] = this._options[key] ? '1' : '0';
        });

        if (this._passwordHash) node._$password = this._passwordHash;
        if (this._algorithmName) node._$algorithmName = this._algorithmName;
        if (this._hashValue) node._$hashValue = this._hashValue;
        if (this._saltValue) node._$saltValue = this._saltValue;
        if (this._spinCount !== null && !Number.isNaN(this._spinCount)) {
            node._$spinCount = String(this._spinCount);
        }

        return node;
    }

    // ==================== 内部方法 ====================

    _setOption(key, value) {
        if (this._options[key] === value) return;
        this._options[key] = value;
        if (this._enabled) {
            this._unregisterListeners();
            this._registerListeners();
        }
    }

    _getListenerId(key) {
        return `protection_${this._sheet.name}_${key}`;
    }

    _registerListeners() {
        if (this._listenersBound) {
            this._unregisterListeners();
        }

        const sheet = this._sheet;
        const SN = this.SN;
        const opts = this._options;
        const t = (key, params) => SN.t(key, params);

        // 1. 单元格编辑保护（锁定单元格不可编辑）
        SN.Event.on('beforeCellEdit', (e) => {
            if (e.data.sheet !== sheet) return;
            if (this.isCellLocked(e.data.cell)) {
                const msg = t('core.sheet.sheetProtection.message.cellProtected');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }
        }, { id: this._getListenerId('cellEdit'), priority: EventPriority.PROTECTION });

        // 2. 插入行保护
        if (!opts.insertRows) {
            SN.Event.on('beforeInsertRows', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.insertRowsDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('insertRows'), priority: EventPriority.PROTECTION });
        }

        // 3. 插入列保护
        if (!opts.insertColumns) {
            SN.Event.on('beforeInsertColumns', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.insertColsDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('insertColumns'), priority: EventPriority.PROTECTION });
        }

        // 4. 删除行保护
        if (!opts.deleteRows) {
            SN.Event.on('beforeDeleteRows', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.deleteRowsDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('deleteRows'), priority: EventPriority.PROTECTION });
        }

        // 5. 删除列保护
        if (!opts.deleteColumns) {
            SN.Event.on('beforeDeleteColumns', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.deleteColsDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('deleteColumns'), priority: EventPriority.PROTECTION });
        }

        // 6. 格式化单元格保护（合并/取消合并）
        if (!opts.formatCells) {
            SN.Event.on('beforeMerge', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.mergeDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('merge'), priority: EventPriority.PROTECTION });

            SN.Event.on('beforeUnmerge', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.unmergeDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('unmerge'), priority: EventPriority.PROTECTION });

            // 样式修改保护
            SN.Event.on('beforeCellStyleChange', (e) => {
                if (e.data.sheet !== sheet) return;
                if (this.isCellLocked(e.data.cell)) {
                    const msg = t('core.sheet.sheetProtection.message.styleDenied');
                    e.cancel(msg);
                    SN.Utils.toast(msg);
                }
            }, { id: this._getListenerId('cellStyle'), priority: EventPriority.PROTECTION });
        }

        // 7. 格式化行保护（行高调整）
        if (!opts.formatRows) {
            SN.Event.on('beforeRowResize', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.rowResizeDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('rowResize'), priority: EventPriority.PROTECTION });

            SN.Event.on('beforeRowHide', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.rowHideDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('rowHide'), priority: EventPriority.PROTECTION });

            SN.Event.on('beforeRowUnhide', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.rowUnhideDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('rowUnhide'), priority: EventPriority.PROTECTION });
        }

        // 8. 格式化列保护（列宽调整）
        if (!opts.formatColumns) {
            SN.Event.on('beforeColResize', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.colResizeDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('colResize'), priority: EventPriority.PROTECTION });

            SN.Event.on('beforeColHide', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.colHideDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('colHide'), priority: EventPriority.PROTECTION });

            SN.Event.on('beforeColUnhide', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.colUnhideDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('colUnhide'), priority: EventPriority.PROTECTION });
        }

        // 9. 插入超链接保护
        if (!opts.insertHyperlinks) {
            SN.Event.on('beforeHyperlinkInsert', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.hyperlinkDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('hyperlink'), priority: EventPriority.PROTECTION });
        }

        // 10. 排序保护
        if (!opts.sort) {
            SN.Event.on('beforeSort', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.sortDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('sort'), priority: EventPriority.PROTECTION });
        }

        // 11. 自动筛选保护
        if (!opts.autoFilter) {
            SN.Event.on('beforeAutoFilter', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.autoFilterDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('autoFilter'), priority: EventPriority.PROTECTION });
        }

        // 12. 数据透视表保护
        if (!opts.pivotTables) {
            SN.Event.on('beforePivotTableChange', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.pivotDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('pivotTable'), priority: EventPriority.PROTECTION });
        }

        // 13. 对象保护（监听 Drawing 的添加和删除）
        if (!opts.objects) {
            SN.Event.on('beforeDrawingAdd', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.addObjectDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('drawingAdd'), priority: EventPriority.PROTECTION });

            SN.Event.on('beforeDrawingRemove', (e) => {
                if (e.data.sheet !== sheet) return;
                const msg = t('core.sheet.sheetProtection.message.removeObjectDenied');
                e.cancel(msg);
                SN.Utils.toast(msg);
            }, { id: this._getListenerId('drawingRemove'), priority: EventPriority.PROTECTION });
        }

        // 14. 选择单元格保护
        const canSelectLocked = opts.selectLockedCells !== false;
        const canSelectUnlocked = opts.selectUnlockedCells !== false;
        if (!canSelectLocked || !canSelectUnlocked) {
            let lastToastAt = 0;
            const toastOnce = (msg) => {
                const now = Date.now();
                if (now - lastToastAt < 500) return;
                lastToastAt = now;
                SN.Utils.toast(msg);
            };

            SN.Event.on('beforeSelectionChange', (e) => {
                if (e.data.sheet !== sheet) return;

                if (e.data.type === 'activeCell' && e.data.newCell) {
                    const cell = sheet.getCell(e.data.newCell.r, e.data.newCell.c);
                    const locked = this.isCellLocked(cell);
                    if (locked && !canSelectLocked) {
                        const msg = t('core.sheet.sheetProtection.message.cannotSelectLocked');
                        e.cancel(msg);
                        toastOnce(msg);
                    } else if (!locked && !canSelectUnlocked) {
                        const msg = t('core.sheet.sheetProtection.message.cannotSelectUnlocked');
                        e.cancel(msg);
                        toastOnce(msg);
                    }
                    return;
                }

                if (e.data.type === 'activeAreas' && e.data.newAreas?.length) {
                    for (const area of e.data.newAreas) {
                        if (!this.isAreaSelectable(area)) {
                            const msg = !canSelectLocked
                                ? t('core.sheet.sheetProtection.message.cannotSelectLocked')
                                : t('core.sheet.sheetProtection.message.cannotSelectUnlocked');
                            e.cancel(msg);
                            toastOnce(msg);
                            break;
                        }
                    }
                }
            }, { id: this._getListenerId('selection'), priority: EventPriority.PROTECTION });
        }

        this._listenersBound = true;
    }

    _unregisterListeners() {
        if (!this._listenersBound) return;

        const SN = this.SN;
        const listenerKeys = [
            'cellEdit', 'insertRows', 'insertColumns', 'deleteRows', 'deleteColumns',
            'merge', 'unmerge', 'cellStyle', 'rowResize', 'rowHide', 'rowUnhide',
            'colResize', 'colHide', 'colUnhide', 'hyperlink', 'sort', 'autoFilter',
            'pivotTable', 'drawingAdd', 'drawingRemove', 'selection'
        ];
        const eventNames = [
            'beforeCellEdit', 'beforeInsertRows', 'beforeInsertColumns', 'beforeDeleteRows', 'beforeDeleteColumns',
            'beforeMerge', 'beforeUnmerge', 'beforeCellStyleChange', 'beforeRowResize', 'beforeRowHide', 'beforeRowUnhide',
            'beforeColResize', 'beforeColHide', 'beforeColUnhide', 'beforeHyperlinkInsert', 'beforeSort', 'beforeAutoFilter',
            'beforePivotTableChange', 'beforeDrawingAdd', 'beforeDrawingRemove', 'beforeSelectionChange'
        ];

        listenerKeys.forEach((key, index) => {
            SN.Event.off(eventNames[index], this._getListenerId(key));
        });

        this._listenersBound = false;
    }
}

// 导出常量供外部使用
export { PROTECTION_KEYS, PROTECTION_DEFAULTS, hashPassword };
