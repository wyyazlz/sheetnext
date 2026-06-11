/**
 * DependencyGraph - Formula Dependency Chain Manager
 *
 * Features:
 * 1. Track dependencies between cells
 * 2. Automatically calculate affected cells
 * 3. Detect circular references
 * 4. Topological ranking ensures correct calculation order
 * 5. Support cross-sheet references
 * 6. Volatile function support (Rand, today, now, etc.)
 *
 * Performance Optimization:
 * - Addition and deletion of O (1) complexity using Map/Set
 * - Incremental updates, only affected cells are recalculated
 * - Batch update mode reduces double counting
 * - Cache topology sort results
 */

export default class DependencyGraph {
    /** @param {import('../Workbook/Workbook.js').default} SN */
    constructor(SN) {
        /** @type {import('../Workbook/Workbook.js').default} */
        this.SN = SN;

        // 核心数据结构
        // dependents: 反向依赖图 - B -> Set[A1, A2, ...] 表示"B被哪些单元格依赖"
        // 当B改变时，需要重新计算A1, A2等
        this.dependents = new Map();

        // dependencies: 正向依赖图 - A -> Set[B1, B2, ...] 表示"A依赖哪些单元格"
        // 用于更新A时清理旧依赖，建立新依赖
        this.dependencies = new Map();

        // Range dependencies are tracked as areas to avoid expanding whole rows/columns.
        this.rangeDependencies = new Map();

        // Volatile 单元格集合 - 包含 RAND, TODAY, NOW 等函数的单元格
        // 这些单元格在每次重算时都需要清除缓存
        this.volatileCells = new Set();

        // 批量更新模式标志
        this.batchMode = false;
        this.batchUpdates = new Set(); // 批量模式下收集的待更新单元格

        // 计算中的单元格栈（用于循环检测）
        this.calculatingStack = [];
    }

    /**
     * Generate Cell Unique Key
     * @param {string} sheetName - Sheet name
     * @param {number} r - Row index
     * @param {number} c - Column Index
     * @ returns {string} unique key in the format: "sheetName!R {r} C {c} "
     */
    cellKey(sheetName, r, c) {
        return `${sheetName}!R${r}C${c}`;
    }

    /**
     * Parse Cell Key
     * @param {string} key - Cell Keys
     * @returns {{sheetName: string, r: number, c: number}}
     */
    parseKey(key) {
        const match = key.match(/^(.+)!R(\d+)C(\d+)$/);
        if (!match) throw new Error(`Invalid cell key: ${key}`);
        return {
            sheetName: match[1],
            r: parseInt(match[2]),
            c: parseInt(match[3])
        };
    }

    /**
     * Register Volatile Cells
     * @param {string} sheetName - Sheet name
     * @param {number} r - Row index
     * @param {number} c - Column Index
     */
    registerVolatile(sheetName, r, c) {
        const key = this.cellKey(sheetName, r, c);
        this.volatileCells.add(key);
    }

    /**
     * Unregister Volatile Cells
     * @param {string} sheetName - Sheet name
     * @param {number} r - Row index
     * @param {number} c - Column Index
     */
    unregisterVolatile(sheetName, r, c) {
        const key = this.cellKey(sheetName, r, c);
        this.volatileCells.delete(key);
    }

    /**
     * Check if the cell is Volatile
     * @param {string} sheetName - Sheet name
     * @param {number} r - Row index
     * @param {number} c - Column Index
     * @returns {boolean}
     */
    isVolatile(sheetName, r, c) {
        const key = this.cellKey(sheetName, r, c);
        return this.volatileCells.has(key);
    }

    _collectVolatileRecalcKeys(seedKeys = []) {
        const out = new Set(seedKeys);
        for (const key of this.volatileCells) {
            out.add(key);
            const affected = this.getAffectedCells(key);
            for (const affectedKey of affected) out.add(affectedKey);
        }
        return out;
    }

    _collectSheetFormulaKeys(sheet, keys = new Set()) {
        if (!sheet) return keys;
        for (let r = 0; r < sheet.rows.length; r++) {
            const row = sheet.rows[r];
            if (!row) continue;
            for (let c = 0; c < row.cells.length; c++) {
                const cell = row.cells[c];
                if (cell?.isFormula) keys.add(this.cellKey(sheet.name, r, c));
            }
        }
        return keys;
    }

    _recalculateKeys(keys) {
        if (!keys.size) return;
        const sorted = this.topologicalSort(keys);
        for (const key of sorted) {
            try {
                const { sheetName, r, c } = this.parseKey(key);
                const sheet = this.SN.getSheet(sheetName);
                const cell = sheet?.getCell(r, c);
                if (!cell) continue;

                cell._calcVal = undefined;
                cell._showVal = undefined;
                cell._fmtResult = undefined;
                const _ = cell.calcVal;
            } catch (e) {
                console.warn(`Failed to recalculate cell ${key}:`, e);
            }
        }
    }

    /**
     * Recalculate all Volatile cells and their dependency chains
     * This is the core approach of the Volatile mechanism
     * When to call: When the user manually triggers a recalculation after editing any cell
     */
    recalculateVolatile() {
        if (this.volatileCells.size === 0) return;
        this._recalculateKeys(this._collectVolatileRecalcKeys());
    }

    /**
     * Recalculate formulas on one sheet.
     * @param {import('../Sheet/Sheet.js').default} sheet - Target sheet
     */
    recalculateSheetFormulas(sheet) {
        this._recalculateKeys(this._collectSheetFormulaKeys(sheet));
    }

    /**
     * Recalculate formulas on all initialized sheets.
     */
    recalculateAllFormulas() {
        const keys = new Set();
        for (const sheet of this.SN.sheets || []) {
            if (!sheet?.initialized) continue;
            this._collectSheetFormulaKeys(sheet, keys);
        }
        this._recalculateKeys(keys);
    }

    /**
     * Get Volatile cell count
     * @returns {number}
     */
    get volatileCount() {
        return this.volatileCells.size;
    }

    /**
     * Add dependency
     * @param {string} fromKey - Dependent Cell Key (Formula Cell)
     * @param {string} toKey - Dependent Party Cell Key (Reference Cell)
     */
    addDependency(fromKey, toKey) {
        // 正向依赖：fromKey依赖toKey
        if (!this.dependencies.has(fromKey)) {
            this.dependencies.set(fromKey, new Set());
        }
        this.dependencies.get(fromKey).add(toKey);

        // 反向依赖：toKey被fromKey依赖
        if (!this.dependents.has(toKey)) {
            this.dependents.set(toKey, new Set());
        }
        this.dependents.get(toKey).add(fromKey);
    }

    /**
     * Remove dependency
     * @param {string} fromKey - Relying Party Cell Key
     * @param {string} toKey - Dependent party cell key (optional, remove all if not passed)
     */
    removeDependency(fromKey, toKey = null) {
        if (toKey) {
            // 移除特定依赖
            this.dependencies.get(fromKey)?.delete(toKey);
            this.dependents.get(toKey)?.delete(fromKey);
        } else {
            // 移除fromKey的所有依赖
            const deps = this.dependencies.get(fromKey);
            if (deps) {
                for (const depKey of deps) {
                    this.dependents.get(depKey)?.delete(fromKey);
                }
                this.dependencies.delete(fromKey);
            }
            this.rangeDependencies.delete(fromKey);
        }
    }

    /**
     * Add an area dependency without expanding every cell in the area.
     * @param {string} fromKey - Formula cell key
     * @param {string} sheetName - Dependency sheet name
     * @param {{s:{r:number,c:number},e:{r:number,c:number}}} area - Dependency area
     */
    addRangeDependency(fromKey, sheetName, area) {
        if (!this.rangeDependencies.has(fromKey)) {
            this.rangeDependencies.set(fromKey, []);
        }
        this.rangeDependencies.get(fromKey).push({
            sheetName,
            s: { r: area.s.r, c: area.s.c },
            e: { r: area.e.r, c: area.e.c }
        });
    }

    /**
     * Update Cell Dependencies
     * Called when the cell formula changes
     * @param {Cell} cell - Cell Object
     */
    updateDependencies(cell) {
        const sheetName = cell.row.sheet.name;
        const fromKey = this.cellKey(sheetName, cell.row.rIndex, cell.cIndex);

        // 先移除旧的依赖关系
        this.removeDependency(fromKey);

        // 如果不是公式，直接返回
        if (!cell.isFormula) return;

        try {
            const formulaStr = cell.editVal.substring(1); // 去掉开头的=
            const areas = this.SN.Formula.parseDependencyAreas(formulaStr, sheetName);

            for (const dep of areas) {
                const depSheetName = dep.sheetName || sheetName;
                if (dep.s.r === dep.e.r && dep.s.c === dep.e.c) {
                    this.addDependency(fromKey, this.cellKey(depSheetName, dep.s.r, dep.s.c));
                } else {
                    this.addRangeDependency(fromKey, depSheetName, dep);
                }
            }
        } catch (error) {
            console.warn(`Failed to parse dependencies for ${fromKey}:`, error);
        }
    }

    _updateDependenciesFromFormula(sheetName, r, c, formulaStr) {
        const fromKey = this.cellKey(sheetName, r, c);
        this.removeDependency(fromKey);
        if (!formulaStr) return;

        try {
            const areas = this.SN.Formula.parseDependencyAreas(formulaStr, sheetName);
            for (const dep of areas) {
                const depSheetName = dep.sheetName || sheetName;
                if (dep.s.r === dep.e.r && dep.s.c === dep.e.c) {
                    this.addDependency(fromKey, this.cellKey(depSheetName, dep.s.r, dep.s.c));
                } else {
                    this.addRangeDependency(fromKey, depSheetName, dep);
                }
            }
        } catch (error) {
            console.warn(`Failed to parse dependencies for ${fromKey}:`, error);
        }
    }

    _rebuildPendingFormulaDependencies(sheet) {
        const formulas = Array.isArray(sheet?._pendingXlsxRowMeta?.formulas)
            ? sheet._pendingXlsxRowMeta.formulas
            : [];
        if (formulas.length === 0) return;

        const sheetName = sheet.name;
        formulas.forEach(item => {
            const r = Math.max(0, Number(item?.r) || 0);
            const c = Math.max(0, Number(item?.c) || 0);
            if (sheet.rows?.[r]) return;
            let formula = item?.f || '';
            if (!formula && item?.si !== undefined) {
                const shared = this.SN._sharedFormula?.[item.si];
                if (shared?.f) {
                    formula = this.SN.Utils.getFillFormula(shared.f, shared.m.r, shared.m.c, r, c);
                }
            }
            this._updateDependenciesFromFormula(sheetName, r, c, formula);
        });
    }

    /**
     * get affected cells (recursively get all dependency chains)
     * @param {string} cellKey - changed cell keys
     * @ returns {Set<string>} set of affected cell keys
     */
    getAffectedCells(cellKey) {
        const affected = new Set();
        const queue = [cellKey];
        const visited = new Set();

        while (queue.length > 0) {
            const current = queue.shift();

            if (visited.has(current)) continue;
            visited.add(current);

            const dependentCells = this.dependents.get(current);
            if (dependentCells) {
                for (const dep of dependentCells) {
                    affected.add(dep);
                    queue.push(dep);
                }
            }

            for (const dep of this._getRangeDependents(current)) {
                affected.add(dep);
                queue.push(dep);
            }
        }

        return affected;
    }

    _getRangeDependents(cellKey) {
        const { sheetName, r, c } = this.parseKey(cellKey);
        const out = [];

        for (const [fromKey, areas] of this.rangeDependencies) {
            for (const area of areas) {
                if (area.sheetName !== sheetName) continue;
                if (r < area.s.r || r > area.e.r || c < area.s.c || c > area.e.c) continue;
                out.push(fromKey);
                break;
            }
        }

        return out;
    }

    /**
     * Topology Sorting - Determine Calculation Order
     * @param {Set<string>} cellKeys - Collection of cell keys that need to be sorted
     * @ returns {Array<string>} Sorted Cell Key Array
     */
    topologicalSort(cellKeys) {
        const sorted = [];
        const visited = new Set();
        const temp = new Set(); // 临时标记（用于检测循环）

        const visit = (key) => {
            if (temp.has(key)) {
                // 检测到循环依赖
                throw new Error(`检测到循环引用: ${key}`);
            }
            if (visited.has(key)) return;

            temp.add(key);

            // 访问所有依赖
            const deps = this.dependencies.get(key);
            if (deps) {
                for (const dep of deps) {
                    // 只访问在cellKeys集合中的依赖
                    if (cellKeys.has(dep)) {
                        visit(dep);
                    }
                }
            }

            temp.delete(key);
            visited.add(key);
            sorted.push(key);
        };

        for (const key of cellKeys) {
            if (!visited.has(key)) {
                try {
                    visit(key);
                } catch (error) {
                    console.error(error.message);
                    // 继续处理其他单元格
                }
            }
        }

        return sorted;
    }

    /**
     * Detection Cycle Reference
     * @param {string} fromKey - starting cell
     * @param {string} toKey - target cell
Is there a cycle for
     * @ returns {boolean}
     */
    detectCycle(fromKey, toKey) {
        if (fromKey === toKey) return true;

        const visited = new Set();
        const queue = [toKey];

        while (queue.length > 0) {
            const current = queue.shift();
            if (current === fromKey) return true;
            if (visited.has(current)) continue;

            visited.add(current);

            const deps = this.dependents.get(current);
            if (deps) {
                for (const dep of deps) {
                    queue.push(dep);
                }
            }
        }

        return false;
    }

    /**
     * Start Batch Update Mode
     */
    startBatch() {
        this.batchMode = true;
        this.batchUpdates.clear();
    }

    /**
     * ends bulk update mode and performs all updates
Trigger update when
     */
    endBatch() {
        this.batchMode = false;

        if (this.batchUpdates.size === 0) return;

        // 拓扑排序确定计算顺序
        const sorted = this.topologicalSort(this.batchUpdates);

        // 按顺序重新计算
        for (const key of sorted) {
            const { sheetName, r, c } = this.parseKey(key);
            const sheet = this.SN.getSheet(sheetName);
            if (sheet) {
                const cell = sheet.getCell(r, c);
                // 清除缓存，强制重新计算
                cell._calcVal = undefined;
                cell._showVal = undefined;
                // 触发计算（通过访问calcVal）
                const _ = cell.calcVal;
            }
        }

        this.batchUpdates.clear();
    }

    /**
     * cell value changes
     * @param {string} sheetName - Sheet name
     * @param {number} r - Row index
     * @param {number} c - Column Index
     */
    notifyChange(sheetName, r, c) {
        const cellKey = this.cellKey(sheetName, r, c);

        const affected = this._collectVolatileRecalcKeys(this.getAffectedCells(cellKey));

        if (affected.size === 0) return;

        if (this.batchMode) {
            // 批量模式：收集待更新的单元格
            for (const key of affected) {
                this.batchUpdates.add(key);
            }
        } else {
            this._recalculateKeys(affected);
        }
    }

    /**
     * Rebuilding dependencies across worksheets
     * @param {Sheet} sheet - Sheet Objects
     */
    rebuildForSheet(sheet) {
        const sheetName = sheet.name;

        // 清除该工作表的所有依赖
        const keysToRemove = [];
        for (const key of this.dependencies.keys()) {
            if (key.startsWith(`${sheetName}!`)) {
                keysToRemove.push(key);
            }
        }
        for (const key of keysToRemove) {
            this.removeDependency(key);
        }

        // 重新扫描所有单元格
        this._rebuildPendingFormulaDependencies(sheet);

        for (let r = 0; r < sheet.rowCount; r++) {
            const row = sheet.rows[r];
            if (!row) continue;

            for (let c = 0; c < row.cells.length; c++) {
                const cell = row.cells[c];
                if (cell && cell.isFormula) {
                    this.updateDependencies(cell);
                }
            }
        }
    }

    /**
     * Rebuild dependencies for all worksheets
     */
    rebuildAll() {
        // 清空所有依赖
        this.dependencies.clear();
        this.dependents.clear();
        this.rangeDependencies.clear();

        // 重建每个工作表
        for (const sheet of this.SN.sheets) {
            if (!sheet.initialized) sheet._init();
            this.rebuildForSheet(sheet);
        }
    }

    /**
     * Get direct dependencies on cells (who I depend on)
     * @param {string} sheetName - Sheet name
     * @param {number} r - Row index
     * @param {number} c - Column Index
     * @ returns {Array} dependent array of cells
     */
    getDependencies(sheetName, r, c) {
        const key = this.cellKey(sheetName, r, c);
        const deps = this.dependencies.get(key);
        if (!deps) return [];

        return Array.from(deps).map(k => this.parseKey(k));
    }

    /**
     * Get the direct dependents of a cell (who depends on me)
     * @param {string} sheetName - Sheet name
     * @param {number} r - Row index
     * @param {number} c - Column Index
     * @ returns {Array} Dependent Cell Array
     */
    getDependents(sheetName, r, c) {
        const key = this.cellKey(sheetName, r, c);
        const deps = this.dependents.get(key);
        if (!deps) return [];

        return Array.from(deps).map(k => this.parseKey(k));
    }

    /**
     * Clear all dependencies
     */
    clear() {
        this.dependencies.clear();
        this.dependents.clear();
        this.rangeDependencies.clear();
        this.volatileCells.clear();
        this.batchUpdates.clear();
    }

    /**
     * Get dependency diagram statistics (for debugging)
     */
    getStats() {
        return {
            totalDependencies: this.dependencies.size,
            totalDependents: this.dependents.size,
            rangeDependencies: this.rangeDependencies.size,
            volatileCells: this.volatileCells.size,
            batchMode: this.batchMode,
            batchUpdatesCount: this.batchUpdates.size
        };
    }
}
