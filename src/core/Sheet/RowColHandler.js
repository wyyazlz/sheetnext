/**
 * 行列操作处理模块
 * 负责处理行列的插入和删除操作
 */

/**
 * 插入行
 * @param {Number} r - 插入位置的行索引
 * @param {Number} number - 插入的行数
 */
export function addRows(r, number = 1) {
    // 触发 beforeInsertRows 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeInsertRows', {
        sheet: this,
        row: r,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('insertRows', {
        sheet: this,
        row: r,
        count: number
    }, () => {
    // 保存受影响的合并单元格信息,用于撤销
    const cds = this.merges.filter(m => m.s.r < r && m.e.r >= r);
    const ams = this.merges.filter(m => m.s.r >= r);
    const oldMerges = [...cds, ...ams].map(m => ({ ...m })); // 深拷贝保存原始状态
    const oldStyledRows = new Set(this._styledRows); // 保存旧的样式行索引
    // 保存受影响的超级表信息（用于撤销）
    const affectedTables = [];
    if (this.Table && this.Table.size > 0) {
        this.Table.forEach(table => {
            const range = table.range;
            if (range && r >= range.s.r && r <= range.e.r + 1) {
                // 插入位置在表格范围内或紧邻底部时，按 Excel 行为扩展表格高度
                affectedTables.push({
                    table,
                    oldRange: { s: { ...range.s }, e: { ...range.e } }
                });
            }
        });
    }

    const applyRowInsertToTables = () => {
        affectedTables.forEach(({ table, oldRange }) => {
            table._ref = {
                s: { ...oldRange.s },
                e: { ...oldRange.e, r: oldRange.e.r + number }
            };
            if (table.autoFilterEnabled) {
                table._updateAutoFilter();
            }
        });
    };
    const restoreTables = () => {
        affectedTables.forEach(({ table, oldRange }) => {
            table._ref = {
                s: { ...oldRange.s },
                e: { ...oldRange.e }
            };
            if (table.autoFilterEnabled) {
                table._updateAutoFilter();
            }
        });
    };

    this.rows.splice(r, 0, ...new Array(number));
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledRows 索引
    const newStyledRows = new Set();
    this._styledRows.forEach(idx => {
        if (idx >= r) newStyledRows.add(idx + number);
        else newStyledRows.add(idx);
    });
    this._styledRows = newStyledRows;

    // 更新受影响的合并单元格
    cds.forEach(m => {
        m.e.r += number;
        for (let i = r; i < r + number; i++) {
            for (let j = m.s.c; j <= m.e.c; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });

    ams.forEach(m => {
        m.s.r += number;
        m.e.r += number;
        for (let i = m.s.r; i <= m.e.r; i++) {
            for (let j = m.s.c; j <= m.e.c; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });
    applyRowInsertToTables();

    // 添加撤销重做操作
    const newStyledRowsAfter = new Set(this._styledRows); // 保存新的样式行索引
    this.SN.UndoRedo.add({
        undo: () => {
            // 删除插入的行
            this.rows.splice(r, number);
            // 恢复样式行索引
            this._styledRows = new Set(oldStyledRows);

            // 恢复合并单元格到原始状态
            [...cds, ...ams].forEach((m, i) => {
                Object.assign(m, oldMerges[i]);
                // 更新单元格的合并状态
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });
            restoreTables();
        },
        redo: () => {
            // 重新插入行
            this.rows.splice(r, 0, ...new Array(number));
            // 恢复样式行索引
            this._styledRows = new Set(newStyledRowsAfter);

            // 更新合并单元格
            cds.forEach(m => {
                m.e.r += number;
                for (let i = r; i < r + number; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            ams.forEach(m => {
                m.s.r += number;
                m.e.r += number;
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });
            applyRowInsertToTables();
        }
    });

    // 调整图纸位置（行插入）
    this.Drawing._onRowsInserted(r, number);

    // 触发 afterInsertRows 事件
    this.SN.Event.emit('afterInsertRows', {
        sheet: this,
        row: r,
        count: number
    });
    this.SN._recordChange({
        type: 'insertRows',
        sheet: this,
        row: r,
        count: number
    });
    });
}

/**
 * 插入列
 * @param {Number} c - 插入位置的列索引
 * @param {Number} number - 插入的列数
 */
export function addCols(c, number = 1) {
    // 触发 beforeInsertColumns 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeInsertColumns', {
        sheet: this,
        col: c,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('insertColumns', {
        sheet: this,
        col: c,
        count: number
    }, () => {
    // 保存受影响的合并单元格信息,用于撤销
    const cdsc = this.merges.filter(m => m.s.c < c && m.e.c >= c);
    const amc = this.merges.filter(m => m.s.c >= c);
    const oldMerges = [...cdsc, ...amc].map(m => ({ ...m })); // 深拷贝保存原始状态
    const oldStyledCols = new Set(this._styledCols); // 保存旧的样式列索引

    // 保存受影响的超级表信息（用于撤销）
    const affectedTables = [];
    if (this.Table && this.Table.size > 0) {
        this.Table.forEach(table => {
            const range = table.range;
            if (range && c >= range.s.c && c <= range.e.c + 1) {
                // 插入位置在表格范围内或紧邻右侧
                affectedTables.push({
                    table,
                    oldRange: { s: { ...range.s }, e: { ...range.e } },
                    oldColumns: table.columns.map(col => ({ ...col }))
                });
            }
        });
    }

    // 保存受影响的自动筛选信息（非超级表）
    let oldAutoFilterRange = null;
    const autoFilter = this.AutoFilter;
    const afRange = autoFilter?.range;
    const autoFilterAffected = afRange && c >= afRange.s.c && c <= afRange.e.c + 1;
    if (autoFilterAffected) {
        oldAutoFilterRange = { s: { ...afRange.s }, e: { ...afRange.e } };
    }

    // 在每一行插入新列
    this.rows.forEach(row => {
        if (row) row.cells.splice(c, 0, ...new Array(number));
    });
    this.cols.splice(c, 0, ...new Array(number));
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledCols 索引
    const newStyledCols = new Set();
    this._styledCols.forEach(idx => {
        if (idx >= c) newStyledCols.add(idx + number);
        else newStyledCols.add(idx);
    });
    this._styledCols = newStyledCols;

    // 更新受影响的合并单元格的结束列
    cdsc.forEach(m => {
        m.e.c += number;
        // 给新单元格打上标记
        for (let i = m.s.r; i <= m.e.r; i++) {
            for (let j = c; j < c + number; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });

    amc.forEach(m => {
        m.s.c += number;
        m.e.c += number;
        for (let i = m.s.r; i <= m.e.r; i++) {
            for (let j = m.s.c; j <= m.e.c; j++) {
                const cell = this.getCell(i, j);
                cell.master = { r: m.s.r, c: m.s.c };
                cell.isMerged = true;
            }
        }
    });

    // 更新受影响的超级表
    affectedTables.forEach(({ table, oldRange }) => {
        const range = table.range;
        const insertColIndex = c - range.s.c; // 相对于表格的插入位置

        // 扩展表格范围
        range.e.c += number;

        // 在插入位置添加新列定义
        for (let i = 0; i < number; i++) {
            table.columns.splice(insertColIndex, 0, {
                id: table.columns.length + i + 1,
                name: `列${insertColIndex + i + 1}`
            });
        }

        // 更新自动筛选范围
        if (table.autoFilterEnabled) {
            table._updateAutoFilter();
        }
    });

    // 更新普通自动筛选范围（非超级表）
    if (autoFilterAffected && afRange) {
        afRange.e.c += number;
        autoFilter._ref = this.SN.Utils.rangeNumToStr(afRange);
    }

    // 添加撤销重做操作
    const newStyledColsAfter = new Set(this._styledCols); // 保存新的样式列索引
    this.SN.UndoRedo.add({
        undo: () => {
            // 删除插入的列
            this.rows.forEach(row => {
                if (row) row.cells.splice(c, number);
            });
            this.cols.splice(c, number);
            // 恢复样式列索引
            this._styledCols = new Set(oldStyledCols);

            // 恢复合并单元格到原始状态
            [...cdsc, ...amc].forEach((m, i) => {
                Object.assign(m, oldMerges[i]);
                // 更新单元格的合并状态
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            // 恢复超级表到原始状态
            affectedTables.forEach(({ table, oldRange, oldColumns }) => {
                table._ref = oldRange;
                table.columns = oldColumns;
                if (table.autoFilterEnabled) {
                    table._updateAutoFilter();
                }
            });

            // 恢复自动筛选范围
            if (oldAutoFilterRange) {
                autoFilter._range = oldAutoFilterRange;
                autoFilter._ref = this.SN.Utils.rangeNumToStr(oldAutoFilterRange);
            }
        },
        redo: () => {
            // 重新插入列
            this.rows.forEach(row => {
                if (row) row.cells.splice(c, 0, ...new Array(number));
            });
            this.cols.splice(c, 0, ...new Array(number));
            // 恢复样式列索引
            this._styledCols = new Set(newStyledColsAfter);

            // 更新合并单元格
            cdsc.forEach(m => {
                m.e.c += number;
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = c; j < c + number; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            amc.forEach(m => {
                m.s.c += number;
                m.e.c += number;
                for (let i = m.s.r; i <= m.e.r; i++) {
                    for (let j = m.s.c; j <= m.e.c; j++) {
                        const cell = this.getCell(i, j);
                        cell.master = { r: m.s.r, c: m.s.c };
                        cell.isMerged = true;
                    }
                }
            });

            // 重新更新超级表
            affectedTables.forEach(({ table, oldRange }) => {
                const range = table.range;
                const insertColIndex = c - oldRange.s.c;
                range.e.c += number;
                for (let i = 0; i < number; i++) {
                    table.columns.splice(insertColIndex, 0, {
                        id: table.columns.length + i + 1,
                        name: `列${insertColIndex + i + 1}`
                    });
                }
                if (table.autoFilterEnabled) {
                    table._updateAutoFilter();
                }
            });

            // 重新更新自动筛选
            if (oldAutoFilterRange) {
                afRange.e.c += number;
                autoFilter._ref = this.SN.Utils.rangeNumToStr(afRange);
            }
        }
    });

    // 调整图纸位置（列插入）
    this.Drawing._onColsInserted(c, number);

    // 触发 afterInsertColumns 事件
    this.SN.Event.emit('afterInsertColumns', {
        sheet: this,
        col: c,
        count: number
    });
    this.SN._recordChange({
        type: 'insertColumns',
        sheet: this,
        col: c,
        count: number
    });
    });
}

/**
 * 删除行
 * @param {Number} r - 删除起始位置的行索引
 * @param {Number} number - 删除的行数
 */
export function delRows(r, number = 1) {
    // 触发 beforeDeleteRows 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeDeleteRows', {
        sheet: this,
        row: r,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('deleteRows', {
        sheet: this,
        row: r,
        count: number
    }, () => {
    // 存储将要删除的行对象
    const deletedRows = this.rows.slice(r, r + number);
    // 存储删除前的合并单元格信息
    const mergesBefore = this.merges.slice();
    const oldStyledRows = new Set(this._styledRows); // 保存旧的样式行索引

    // 检测并更新涉及合并单元格
    this.merges = this.merges.filter(m => {
        // 如果删除区域完全在合并单元格之前，无需修改
        if (m.e.r < r) return true;
        // 如果删除区域完全在合并单元格之后，只需更新索引
        if (m.s.r >= r + number) {
            this.eachCells(m, (rowIndex, c) => {
                this.getCell(rowIndex, c).master = { r: m.s.r - number, c: m.s.c }
            })
            m.s.r -= number;
            m.e.r -= number;
            return true;
        }
        // 判断是否需要调整合并单元格的起始或结束行
        if (m.s.r == r) { // 如果删除的是起始行
            if (m.e.r - m.s.r < number) {
                // 如果删除的行数大于等于合并单元格的总行数，该合并单元格应被移除
                return false;
            } else {
                // 更新剩下的单元格的master行
                m.e.r -= number;
                return true;
            }
        } else { // 删除的是合并单元格中间或尾部的行
            m.e.r -= number;
            return true;
        }
    });

    // 删除指定的行
    this.rows.splice(r, number);
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledRows 索引
    const newStyledRows = new Set();
    this._styledRows.forEach(idx => {
        if (idx >= r && idx < r + number) return; // 删除的行，跳过
        if (idx >= r + number) newStyledRows.add(idx - number);
        else newStyledRows.add(idx);
    });
    this._styledRows = newStyledRows;

    // 如果有合并区域剩一个单元格，清除所有
    this.merges = this.merges.filter(m => {
        if (m.s.r == m.e.r && m.s.c == m.e.c) {
            const cell = this.getCell(m.s.r, m.s.c)
            cell.isMerged = false
            cell.master = null
            return false
        } else {
            return true
        }
    })

    // 存储删除后的合并单元格信息
    const mergesAfter = this.merges.slice();
    const newStyledRowsAfter = new Set(this._styledRows); // 保存新的样式行索引

    // 添加撤销重做功能
    this.SN.UndoRedo.add({
        undo: () => {
            // 撤销操作
            // 将删除的行插回原位置
            this.rows.splice(r, 0, ...deletedRows);
            // 恢复样式行索引
            this._styledRows = new Set(oldStyledRows);
            // 恢复删除前的合并单元格信息
            this.merges = mergesBefore.slice();
        },
        redo: () => {
            // 重做操作
            // 再次删除指定的行
            this.rows.splice(r, number);
            // 恢复样式行索引
            this._styledRows = new Set(newStyledRowsAfter);
            // 恢复删除后的合并单元格信息
            this.merges = mergesAfter.slice();
        }
    });

    // 调整图纸位置（行删除）
    this.Drawing._onRowsDeleted(r, number);

    // 触发 afterDeleteRows 事件
    this.SN.Event.emit('afterDeleteRows', {
        sheet: this,
        row: r,
        count: number
    });
    this.SN._recordChange({
        type: 'deleteRows',
        sheet: this,
        row: r,
        count: number
    });
    });
}

/**
 * 删除列
 * @param {Number} c - 删除起始位置的列索引
 * @param {Number} number - 删除的列数
 */
export function delCols(c, number = 1) {
    // 触发 beforeDeleteColumns 事件（可取消）
    const beforeEvent = this.SN.Event.emit('beforeDeleteColumns', {
        sheet: this,
        col: c,
        count: number
    });
    if (beforeEvent.canceled) return;

    this.SN._withOperation('deleteColumns', {
        sheet: this,
        col: c,
        count: number
    }, () => {
    // 存储删除前的状态
    const deletedCols = this.cols.slice(c, c + number);
    const deletedCells = this.rows.map(row => row ? row.cells.slice(c, c + number) : null);
    const mergesBefore = this.merges.slice();
    const oldStyledCols = new Set(this._styledCols); // 保存旧的样式列索引

    // 检测并更新涉及合并单元格
    this.merges = this.merges.filter(m => {
        // 如果删除区域完全在合并单元格之前，无需修改
        if (m.e.c < c) return true;
        // 如果删除区域完全在合并单元格之后，只需更新索引
        if (m.s.c >= c + number) {
            this.eachCells(m, (r, c) => {
                this.getCell(r, c).master = { r: m.s.r, c: m.s.c - number }
            })
            m.s.c -= number;
            m.e.c -= number;
            return true;
        }
        if (m.s.c == c) { // 删的如果是master
            if (m.e.c - m.s.c < number) {
                return false
            } else { // 更新剩下的单元格的master
                m.e.c = m.e.c - number
                return true
            }
        } else { // 删非master
            m.e.c = m.e.c - number
            return true
        }
    });

    // 删除指定的列
    this.rows.forEach(row => {
        if (row) row.cells.splice(c, number);
    });
    this.cols.splice(c, number);
    this._clearSizeCache(); // 清除尺寸缓存

    // 更新 _styledCols 索引
    const newStyledCols = new Set();
    this._styledCols.forEach(idx => {
        if (idx >= c && idx < c + number) return; // 删除的列，跳过
        if (idx >= c + number) newStyledCols.add(idx - number);
        else newStyledCols.add(idx);
    });
    this._styledCols = newStyledCols;

    // 如果剩一个单元格，清除所有
    this.merges = this.merges.filter(m => {
        if (m.s.r == m.e.r && m.s.c == m.e.c) {
            const cell = this.getCell(m.s.r, m.s.c)
            cell.isMerged = false
            cell.master = null
            return false
        } else {
            return true
        }
    })

    // 存储删除后的状态
    const mergesAfter = this.merges.slice();
    const newStyledColsAfter = new Set(this._styledCols); // 保存新的样式列索引

    // 添加撤销重做功能
    this.SN.UndoRedo.add({
        undo: () => {
            // 恢复列
            this.cols.splice(c, 0, ...deletedCols);
            // 恢复单元格数据
            this.rows.forEach((row, rowIndex) => {
                if (row && deletedCells[rowIndex]) {
                    row.cells.splice(c, 0, ...deletedCells[rowIndex]);
                }
            });
            // 恢复样式列索引
            this._styledCols = new Set(oldStyledCols);
            // 恢复合并单元格信息
            this.merges = mergesBefore.slice();
        },
        redo: () => {
            // 重新删除列
            this.rows.forEach(row => {
                if (row) row.cells.splice(c, number);
            });
            this.cols.splice(c, number);
            // 恢复样式列索引
            this._styledCols = new Set(newStyledColsAfter);
            // 恢复删除后的合并单元格信息
            this.merges = mergesAfter.slice();
        }
    });

    // 调整图纸位置（列删除）
    this.Drawing._onColsDeleted(c, number);

    // 触发 afterDeleteColumns 事件
    this.SN.Event.emit('afterDeleteColumns', {
        sheet: this,
        col: c,
        count: number
    });
    this.SN._recordChange({
        type: 'deleteColumns',
        sheet: this,
        col: c,
        count: number
    });
    });
}
