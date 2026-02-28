/**
 * 合并单元格处理模块
 * 负责处理单元格的合并、解除合并及相关操作
 */

import { findMaxCompositeArea } from './helpers.js';

/**
 * 获取区域在此表能拓展的最大区域，根据合并的单元格拓展
 * @param {Object} selection - 选区
 * @returns {Object} 拓展后的最大区域
 */
export function _getExpandMaxArea(selection) {
    if (!this.merges) return selection;
    let maxArea = { ...selection };
    let condition = true
    while (condition) {
        condition = false
        this.merges.forEach(merge => {
            const newArea = findMaxCompositeArea(merge, maxArea);
            if (JSON.stringify(newArea) != JSON.stringify(maxArea)) condition = true
            maxArea = newArea
        });
    }
    return maxArea; // 返回更新后的最大区域
}

/**
 * 检测区域中是否存在合并的单元格
 * @param {Object} area - 区域对象
 * @returns {Boolean} 是否存在合并单元格
 */
export function areaHaveMerge(area) {
    if (area) {
        return this.merges.some(mareaM => !(mareaM.s.c > area.e.c || mareaM.e.c < area.s.c || mareaM.s.r > area.e.r || mareaM.e.r < area.s.r));
    } else {
        return this.merges.length > 0
    }
}

/**
 * 解除合并，传入单元格地址或区域，默认取 activeAreas
 * @param {Object|String|null} cellAd - 单元格地址或区域
 */
export function unMergeCells(cellAd = null) {
    const sheet = this;

    // 触发 beforeUnmerge 事件（可取消）
    const beforeEvent = sheet.SN.Event.emit('beforeUnmerge', {
        sheet,
        cellAd
    });
    if (beforeEvent.canceled) return;

    return sheet.SN._withOperation('unmerge', { sheet, cellAd }, () => {
    // 默认取 activeAreas，找出所有合并的主单元格
    if (!cellAd) {
        const masters = new Set();
        sheet.activeAreas.forEach(area => {
            sheet.eachCells(area, (r, c) => {
                const cell = sheet.getCell(r, c);
                if (cell.master && cell.master.r === r && cell.master.c === c) {
                    masters.add(`${r},${c}`);
                }
            });
        });
        masters.forEach(key => {
            const [r, c] = key.split(',').map(Number);
            unMergeOne.call(sheet, { r, c });
        });

        sheet.SN._recordChange({
            type: 'unmerge',
            sheet,
            areas: sheet.activeAreas
        });
        // 触发 afterUnmerge 事件
        sheet.SN.Event.emit('afterUnmerge', { sheet, cellAd });
        return;
    }

    if (typeof cellAd === 'string') cellAd = sheet.Utils.cellStrToNum(cellAd);
    unMergeOne.call(sheet, cellAd);

    sheet.SN._recordChange({
        type: 'unmerge',
        sheet,
        cellAd
    });
    // 触发 afterUnmerge 事件
    sheet.SN.Event.emit('afterUnmerge', { sheet, cellAd });
    });
}

/**
 * 解除单个合并区域（内部方法）
 */
function unMergeOne(cellAd) {
    const sheet = this;
    const cell = sheet.getCell(cellAd.r, cellAd.c);
    if (!cell.isMerged) return;

    const index = sheet.merges.findIndex(item => item.s.r === cell.master.r && item.s.c === cell.master.c);
    if (index === -1) return;
    const mergeArea = sheet.merges[index];

    // 记录操作前的状态
    const oldMerge = JSON.parse(JSON.stringify(mergeArea));
    const oldCells = [];
    sheet.eachCells(mergeArea, (r, c) => {
        const cell = sheet.getCell(r, c);
        oldCells.push({
            r, c,
            isMerged: cell.isMerged,
            master: cell.master ? { ...cell.master } : null
        });
    });

    // 执行解除合并
    sheet.eachCells(mergeArea, (r, c) => {
        const cell = sheet.getCell(r, c);
        cell.isMerged = false;
        cell.master = null;
    });
    sheet.merges.splice(index, 1);

    // 撤销重做
    sheet.SN.UndoRedo.add({
        undo: () => {
            sheet.eachCells(oldMerge, (r, c) => {
                const cell = sheet.getCell(r, c);
                cell.isMerged = true;
                cell.master = { r: oldMerge.s.r, c: oldMerge.s.c };
            });
            sheet.merges.push(oldMerge);
        },
        redo: () => {
            sheet.eachCells(oldMerge, (r, c) => {
                const cell = sheet.getCell(r, c);
                cell.isMerged = false;
                cell.master = null;
            });
            const idx = sheet.merges.findIndex(item => item.s.r === oldMerge.s.r && item.s.c === oldMerge.s.c);
            if (idx !== -1) sheet.merges.splice(idx, 1);
        }
    });
}

/**
 * 合并单元格，支持多种模式
 * @param {Array|Object|String|null} areas - 区域，默认取 activeAreas
 * @param {String} mode - 模式: 'default'|'center'|'content'|'same'
 */
export function mergeCells(areas = null, mode = 'default') {
    const sheet = this;

    // 触发 beforeMerge 事件（可取消）
    const beforeEvent = sheet.SN.Event.emit('beforeMerge', {
        sheet,
        areas,
        mode
    });
    if (beforeEvent.canceled) return;

    return sheet.SN._withOperation('merge', { sheet, areas, mode }, () => {
    // 默认取当前选中区域
    if (!areas) areas = sheet.activeAreas;
    if (!areas || areas.length === 0) return;

    // 统一转为数组
    if (!Array.isArray(areas)) areas = [areas];

    // 字符串转对象
    areas = areas.map(area => {
        if (typeof area === 'string') return sheet.rangeStrToNum(area);
        return JSON.parse(JSON.stringify(area));
    });

    // same 模式：合并相同内容的相邻单元格（按列检测）
    if (mode === 'same') {
        areas.forEach(area => {
            for (let c = area.s.c; c <= area.e.c; c++) {
                let startR = area.s.r;
                let currentVal = sheet.getCell(startR, c).showVal;

                for (let r = area.s.r + 1; r <= area.e.r + 1; r++) {
                    const val = r <= area.e.r ? sheet.getCell(r, c).showVal : null;
                    if (val !== currentVal || r > area.e.r) {
                        // 值变化或到末尾，合并之前的区域
                        if (r - 1 > startR) {
                            mergeOneArea.call(sheet, { s: { r: startR, c }, e: { r: r - 1, c } }, false);
                        }
                        startR = r;
                        currentVal = val;
                    }
                }
            }
        });

        sheet.SN._recordChange({
            type: 'merge',
            sheet,
            areas,
            mode
        });
        // 触发 afterMerge 事件
        sheet.SN.Event.emit('afterMerge', { sheet, areas, mode });
        return;
    }

    // 处理每个区域
    areas.forEach(area => {
        if (area.s.r === area.e.r && area.s.c === area.e.c) return;

        if (mode === 'content') {
            // 收集所有内容
            const contents = [];
            sheet.eachCells(area, (r, c) => {
                const val = sheet.getCell(r, c).showVal;
                if (val !== '' && val != null) contents.push(val);
            });
            mergeOneArea.call(sheet, area, false);
            if (contents.length > 0) {
                sheet.getCell(area.s.r, area.s.c).editVal = contents.join(' ');
            }
        } else {
            mergeOneArea.call(sheet, area, mode === 'center');
        }
    });

    sheet.SN._recordChange({
        type: 'merge',
        sheet,
        areas,
        mode
    });
    // 触发 afterMerge 事件
    sheet.SN.Event.emit('afterMerge', { sheet, areas, mode });
    });
}

/**
 * 合并单个区域（内部方法）
 */
function mergeOneArea(areaAd, applyCenter = false) {
    const sheet = this;
    if (areaAd.s.r === areaAd.e.r && areaAd.s.c === areaAd.e.c) return;

    // 收集所有要解除的合并区域
    const mergeAreasToRemove = [];
    const seenMasters = new Set();

    sheet.eachCells(areaAd, (r, c) => {
        const cell = sheet.getCell(r, c);
        if (cell.isMerged && cell.master) {
            const masterKey = `${cell.master.r},${cell.master.c}`;
            if (!seenMasters.has(masterKey)) {
                seenMasters.add(masterKey);
                const mergeArea = sheet.merges.find(item =>
                    item.s.r === cell.master.r && item.s.c === cell.master.c
                );
                if (mergeArea) {
                    mergeAreasToRemove.push(JSON.parse(JSON.stringify(mergeArea)));
                }
            }
        }
    });

    // 记录操作前的状态
    const oldMerges = JSON.parse(JSON.stringify(sheet.merges));
    const oldCells = [];
    const oldStyle = sheet.getCell(areaAd.s.r, areaAd.s.c)._style
        ? JSON.parse(JSON.stringify(sheet.getCell(areaAd.s.r, areaAd.s.c)._style))
        : null;

    const affectedAreas = [...mergeAreasToRemove, areaAd];
    affectedAreas.forEach(area => {
        sheet.eachCells(area, (r, c) => {
            if (!oldCells.find(item => item.r === r && item.c === c)) {
                const cell = sheet.getCell(r, c);
                oldCells.push({
                    r, c,
                    isMerged: cell.isMerged,
                    master: cell.master ? { ...cell.master } : null
                });
            }
        });
    });

    // 先解除所有要移除的合并区域
    mergeAreasToRemove.forEach(mergeArea => {
        sheet.eachCells(mergeArea, (r, c) => {
            const cell = sheet.getCell(r, c);
            cell.isMerged = false;
            cell.master = null;
        });
        const index = sheet.merges.findIndex(item =>
            item.s.r === mergeArea.s.r && item.s.c === mergeArea.s.c
        );
        if (index !== -1) sheet.merges.splice(index, 1);
    });

    // 执行新的合并
    sheet.eachCells(areaAd, (r, c) => {
        const cell = sheet.getCell(r, c);
        cell.isMerged = true;
        cell.master = { r: areaAd.s.r, c: areaAd.s.c };
    });
    sheet.merges.push(areaAd);

    // 居中对齐
    if (applyCenter) {
        const masterCell = sheet.getCell(areaAd.s.r, areaAd.s.c);
        masterCell.alignment = { horizontal: 'center', vertical: 'center' };
    }

    // 撤销重做
    sheet.SN.UndoRedo.add({
        undo: () => {
            sheet.merges = JSON.parse(JSON.stringify(oldMerges));
            oldCells.forEach(item => {
                const cell = sheet.getCell(item.r, item.c);
                cell.isMerged = item.isMerged;
                cell.master = item.master;
            });
            if (applyCenter) {
                sheet.getCell(areaAd.s.r, areaAd.s.c)._style = oldStyle;
            }
        },
        redo: () => {
            mergeAreasToRemove.forEach(mergeArea => {
                sheet.eachCells(mergeArea, (r, c) => {
                    const cell = sheet.getCell(r, c);
                    cell.isMerged = false;
                    cell.master = null;
                });
                const index = sheet.merges.findIndex(item =>
                    item.s.r === mergeArea.s.r && item.s.c === mergeArea.s.c
                );
                if (index !== -1) sheet.merges.splice(index, 1);
            });
            sheet.eachCells(areaAd, (r, c) => {
                const cell = sheet.getCell(r, c);
                cell.isMerged = true;
                cell.master = { r: areaAd.s.r, c: areaAd.s.c };
            });
            sheet.merges.push(areaAd);
            if (applyCenter) {
                sheet.getCell(areaAd.s.r, areaAd.s.c).alignment = { horizontal: 'center', vertical: 'center' };
            }
        }
    });
}
