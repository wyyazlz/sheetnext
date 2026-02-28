/**
 * Table 表格管理器
 * 管理工作表中的所有表格
 */
import TableItem from './TableItem.js';
import { generateTableName, toRef } from './TableUtils.js';

export default class Table {
    /**
     * @param {Sheet} sheet
     */
    constructor(sheet) {
        /** @type {Sheet} */
        this.sheet = sheet;
        this._tables = new Map();
    }

    /** 表格数量 */
    get size() {
        return this._tables.size;
    }

    /** @returns {TableItem[]} */
    getAll() {
        return Array.from(this._tables.values());
    }

    /**
     * 根据ID获取表格
     * @param {string} id
     * @returns {TableItem|null}
     */
    get(id) {
        return this._tables.get(id) || null;
    }

    /**
     * 根据名称获取表格
     * @param {string} name
     * @returns {TableItem|null}
     */
    getByName(name) {
        for (const table of this._tables.values()) {
            if (table.name === name || table.displayName === name) {
                return table;
            }
        }
        return null;
    }

    /**
     * 获取包含指定单元格的表格
     * @param {number} row
     * @param {number} col
     * @returns {TableItem|null}
     */
    getTableAt(row, col) {
        for (const table of this._tables.values()) {
            if (table.contains(row, col)) {
                return table;
            }
        }
        return null;
    }

    /**
     * 添加/创建新表格
     * @param {Object} options
     * @returns {TableItem}
     */
    add(options = {}) {
        // 生成唯一名称
        if (!options.name) {
            options.name = generateTableName(this._tables);
        }
        if (!options.displayName) {
            options.displayName = options.name;
        }

        if (this.sheet.SN.Event.hasListeners('beforeTableAdd')) {
            const beforeEvent = this.sheet.SN.Event.emit('beforeTableAdd', {
                sheet: this.sheet,
                options
            });
            if (beforeEvent.canceled) return null;
        }
        const table = new TableItem(this.sheet, options);
        this._tables.set(table.id, table);

        // 判断是否是导入（仅 XML/文件导入路径显式标记）
        const isImport = options._fromImport === true;

        // 同步列名（仅非导入时，导入时列名已从 XML 解析）
        if (!isImport && table.showHeaderRow && table.range) {
            table.syncColumnNames();
        }

        if (!isImport && table.showTotalsRow && table.range) {
            table.applyTotalsRow({ initializeDefaults: true, force: true });
        }

        // 初始化自动筛选（导入时静默同步，避免 vi 未初始化触发渲染）
        if (table.autoFilterEnabled && table.range) {
            this.sheet.AutoFilter.registerTableScope(table, {
                keepFilters: false,
                silent: isImport,
                autoFilterXml: options._tableAutoFilterXml || null,
                restoreHiddenRowsFromSheet: isImport
            });
        }

        if (this.sheet.SN.Event.hasListeners('afterTableAdd')) {
            this.sheet.SN.Event.emit('afterTableAdd', { sheet: this.sheet, table, options });
        }
        return table;
    }

    /**
     * 从选区创建表格
     * @param {Object} area - { s: {r, c}, e: {r, c} }
     * @param {Object} options
     * @returns {TableItem}
     */
    createFromSelection(area, options = {}) {
        const targetArea = {
            s: {
                r: Math.min(area.s.r, area.e.r),
                c: Math.min(area.s.c, area.e.c)
            },
            e: {
                r: Math.max(area.s.r, area.e.r),
                c: Math.max(area.s.c, area.e.c)
            }
        };
        const isEmptyHeaderValue = value => value === undefined || value === null || value.toString().trim() === '';

        // 检查选区是否与现有表格重叠
        for (const table of this._tables.values()) {
            if (this._isOverlapping(targetArea, table.range)) {
                throw new Error('选区与现有表格重叠');
            }
        }

        // 创建列定义
        const colCount = targetArea.e.c - targetArea.s.c + 1;
        const columns = [];
        const hasHeader = options.showHeaderRow !== false;

        for (let i = 0; i < colCount; i++) {
            const col = targetArea.s.c + i;
            const defaultName = `列${i + 1}`;
            let name = defaultName;

            // 如果有表头，从单元格读取列名
            if (hasHeader) {
                const cell = this.sheet.getCell(targetArea.s.r, col);
                const headerValue = cell.calcVal ?? cell.v ?? cell.editVal;
                if (!isEmptyHeaderValue(headerValue)) {
                    name = headerValue.toString();
                } else {
                    // Keep header cells and table metadata consistent for xlsx table export.
                    cell.editVal = defaultName;
                }
            }

            columns.push({ id: i + 1, name });
        }

        return this.add({
            ...options,
            rangeRef: toRef(targetArea),
            columns,
            showHeaderRow: hasHeader
        });
    }

    /**
     * 删除表格
     * @param {string} id - 表格ID
     * @returns {boolean}
     */
    remove(id) {
        const table = this._tables.get(id);
        if (table) {
            if (this.sheet.SN.Event.hasListeners('beforeTableRemove')) {
                const beforeEvent = this.sheet.SN.Event.emit('beforeTableRemove', { sheet: this.sheet, table, id });
                if (beforeEvent.canceled) return false;
            }
            // 清除筛选
            if (table.autoFilterEnabled) {
                this.sheet.AutoFilter.unregisterTableScope(table.id);
            }
            const removed = this._tables.delete(id);
            if (removed && this.sheet.SN.Event.hasListeners('afterTableRemove')) {
                this.sheet.SN.Event.emit('afterTableRemove', { sheet: this.sheet, table, id });
            }
            return removed;
        }
        return false;
    }

    /**
     * 根据名称删除表格
     * @param {string} name
     * @returns {boolean}
     */
    removeByName(name) {
        const table = this.getByName(name);
        if (table) {
            return this.remove(table.id);
        }
        return false;
    }

    /**
     * 更新名称索引（内部方法，由 TableItem.rename 调用）
     * @param {string} oldName
     * @param {string} newName
     * @param {TableItem} table
     * @private
     */
    _updateIndex(oldName, newName, table) {
        // 当前实现使用 id 作为 key，所以只需确保没有名称冲突
        // 如果将来改为名称索引，这里需要更新 Map
    }

    /**
     * 检查两个范围是否重叠
     * @private
     */
    _isOverlapping(range1, range2) {
        if (!range1 || !range2) return false;
        return !(range1.e.r < range2.s.r || range1.s.r > range2.e.r ||
                 range1.e.c < range2.s.c || range1.s.c > range2.e.c);
    }

    /**
     * 遍历所有表格
     * @param {Function} callback
     */
    forEach(callback) {
        this._tables.forEach(callback);
    }

    /**
     * 转为 JSON 数组
     */
    toJSON() {
        return this.getAll().map(table => table.toJSON());
    }

    /**
     * 从 JSON 恢复
     * @param {Array} jsonArray
     */
    fromJSON(jsonArray) {
        // 清空现有表格
        this._tables.forEach(table => {
            if (table.autoFilterEnabled) {
                this.sheet.AutoFilter.unregisterTableScope(table.id, { silent: true });
            }
        });
        this._tables.clear();

        if (Array.isArray(jsonArray)) {
            jsonArray.forEach(json => {
                const table = TableItem.fromJSON(this.sheet, json);
                this._tables.set(table.id, table);
                if (table.autoFilterEnabled && table.range) {
                    this.sheet.AutoFilter.registerTableScope(table, { keepFilters: false, silent: true });
                }
            });
        }
    }
}
