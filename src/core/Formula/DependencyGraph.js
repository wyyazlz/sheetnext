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

    /**
     * Recalculate all Volatile cells and their dependency chains
     * This is the core approach of the Volatile mechanism
     * When to call: When the user manually triggers a recalculation after editing any cell
     */
    recalculateVolatile() {
        if (this.volatileCells.size === 0) return;

        // 收集所有需要重算的单元格（Volatile + 其依赖链）
        const toRecalc = new Set();

        // 1. 首先清除所有 Volatile 单元格的缓存
        for (const key of this.volatileCells) {
            toRecalc.add(key);

            // 2. 获取依赖 Volatile 单元格的所有单元格
            const affected = this.getAffectedCells(key);
            for (const affectedKey of affected) {
                toRecalc.add(affectedKey);
            }
        }

        // 3. 拓扑排序确定正确的计算顺序
        const sorted = this.topologicalSort(toRecalc);

        // 4. 按顺序清除缓存并重新计算
        for (const key of sorted) {
            try {
                const { sheetName, r, c } = this.parseKey(key);
                const sheet = this.SN.getSheet(sheetName);
                if (sheet) {
                    const cell = sheet.getCell(r, c);
                    if (cell) {
                        // 清除缓存
                        cell._calcVal = undefined;
                        cell._showVal = undefined;
                        cell._fmtResult = undefined;
                        // 触发重新计算
                        const _ = cell.calcVal;
                    }
                }
            } catch (e) {
                console.warn(`Failed to recalculate volatile cell ${key}:`, e);
            }
        }
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
        }
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
            // 使用Formula类的tokenize和parseDeps解析依赖
            const formulaStr = cell.editVal.substring(1); // 去掉开头的=
            const tokens = this.SN.Formula.tokenize(formulaStr);
            const deps = this.SN.Formula.parseDeps(tokens);

            // 添加新的依赖关系
            for (const dep of deps) {
                // 解析是否跨表引用
                let depSheetName = sheetName; // 默认同表
                let depR = dep.r;
                let depC = dep.c;

                // 检查token中是否有跨表引用
                for (const token of tokens) {
                    if (token.type === 'CELL_REF' && token.value.includes('!')) {
                        const [refSheet, refCell] = token.value.split('!');
                        const refSheetClean = refSheet.replace(/^'|'$/g, ''); // 去除引号
                        const refPos = this.SN.Utils.cellStrToNum(refCell);
                        if (refPos.r === depR && refPos.c === depC) {
                            depSheetName = refSheetClean;
                            break;
                        }
                    }
                }

                const toKey = this.cellKey(depSheetName, depR, depC);
                this.addDependency(fromKey, toKey);
            }
        } catch (error) {
            console.warn(`Failed to parse dependencies for ${fromKey}:`, error);
        }
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
        }

        return affected;
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

        // 获取所有受影响的单元格
        const affected = this.getAffectedCells(cellKey);

        if (affected.size === 0) return;

        if (this.batchMode) {
            // 批量模式：收集待更新的单元格
            for (const key of affected) {
                this.batchUpdates.add(key);
            }
        } else {
            // 立即模式：立即更新
            const sorted = this.topologicalSort(affected);

            for (const key of sorted) {
                const { sheetName: s, r: row, c: col } = this.parseKey(key);
                const sheet = this.SN.getSheet(s);
                if (sheet) {
                    const cell = sheet.getCell(row, col);
                    // 清除缓存，强制重新计算
                    cell._calcVal = undefined;
                    cell._showVal = undefined;
                    // 触发计算
                    const _ = cell.calcVal;
                }
            }
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
            volatileCells: this.volatileCells.size,
            batchMode: this.batchMode,
            batchUpdatesCount: this.batchUpdates.size
        };
    }
}
