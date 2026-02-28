/**
 * 视图计算模块
 * 负责计算单元格和区域在可见视图中的位置信息，以及更新视图索引
 */

/**
 * 获取单元格在可见视图中的信息
 * @param {Number} rowIndex - 行索引
 * @param {Number} colIndex - 列索引
 * @param {Boolean} posMerge - 是否处理合并单元格
 * @returns {Object} 包含位置和尺寸信息的对象
 */
export function getCellInViewInfo(rowIndex, colIndex, posMerge = true) {

    let r = rowIndex, c = colIndex, merge;
    if (posMerge) {
        const cell = this.getCell(rowIndex, colIndex);
        if (cell.master) {
            r = cell.master.r;
            c = cell.master.c;
            merge = this.merges.find(m => m.s.r === r && m.s.c === c);
        }
    }

    const viewCol = this.vi.colsIndex
    const viewRow = this.vi.rowsIndex

    // 计算单元格的x坐标
    let cellX = this.indexWidth;
    for (let index of viewCol) {
        if (index >= c) break
        cellX += this.getCol(index).width
    }

    // 计算单元格的y坐标
    let cellY = this.headHeight;
    for (let index of viewRow) {
        if (index >= r) break
        cellY += this.getRow(index).height
    }

    // 获取单元格的宽度和高度
    const cellWidth = this.getCol(colIndex)?.width;
    const cellHeight = this.getRow(rowIndex)?.height;

    // 如果是合并单元格，则重新计算宽度和高度
    const finalWidth = merge ? [...Array(merge.e.c - merge.s.c + 1).keys()].reduce((acc, val) => {
        const col = this.getCol(merge.s.c + val);
        return col.hidden || !this.vi.colsIndex.includes(merge.s.c + val) ? acc : acc + col.width;
    }, 0) : cellWidth;

    const finalHeight = merge ? [...Array(merge.e.r - merge.s.r + 1).keys()].reduce((acc, val) => {
        const row = this.getRow(merge.s.r + val);
        return row.hidden || !this.vi.rowsIndex.includes(merge.s.r + val) ? acc : acc + row.height;
    }, 0) : cellHeight;

    // 判断单元格是否在视图中
    const canvas = this.Canvas;
    const viewWidth = canvas?.viewWidth ?? canvas?.width ?? this._handleLayer.offsetWidth;
    const viewHeight = canvas?.viewHeight ?? canvas?.height ?? this._handleLayer.offsetHeight;
    let inView = false
    if (merge) {
        // 检查矩形的右边和底边是否至少部分进入视图
        let isVisible = (cellX + finalWidth > this.indexWidth) && (cellY + finalHeight > this.headHeight);
        // 检查矩形的左边和顶边是否不完全超出视图
        inView = isVisible && (cellX < viewWidth) && (cellY < viewHeight);
    } else {
        inView = rowIndex >= viewRow[0] && rowIndex <= viewRow[viewRow.length - 1] && colIndex >= viewCol[0] && colIndex <= viewCol[viewCol.length - 1];
    }

    return {
        inView,
        x: cellX,
        y: cellY,
        w: finalWidth,
        h: finalHeight
    };

}

/**
 * 获取区域在可见视图中的信息
 * @param {Object} area - 区域对象
 * @returns {Object} 包含位置和尺寸信息的对象
 */
export function getAreaInviewInfo(area) {

    let { s: { c: startCol, r: startRow }, e: { c: endCol, r: endRow } } = area;

    const viewCol = this.vi.colsIndex
    const viewRow = this.vi.rowsIndex

    // 检查区域是否在可见区域内
    if (endRow < viewRow[0] || startRow > viewRow[viewRow.length - 1] || endCol < viewCol[0] || startCol > viewCol[viewCol.length - 1]) return { inView: false }

    // 使用视图信息和滚动位置来确定区域的起始点
    let startX = this.indexWidth;
    let startY = this.headHeight;
    let totalWidth = 0;
    let totalHeight = 0;

    // 更新区域边界
    if (endRow > viewRow[viewRow.length - 1]) endRow = viewRow[viewRow.length - 1];
    if (endCol > viewCol[viewCol.length - 1]) endCol = viewCol[viewCol.length - 1];

    // 计算起始y坐标和总高度
    for (let index of viewRow) {
        if (index >= startRow && index <= endRow) {
            totalHeight += this.getRow(index).height
        } else if (index < startRow) {
            startY += this.getRow(index).height
        } else {
            break
        }
    }

    // 计算起始x坐标和总宽度
    for (let index of viewCol) {
        if (index >= startCol && index <= endCol) {
            totalWidth += this.getCol(index).width
        } else if (index < startCol) {
            startX += this.getCol(index).width
        } else {
            break
        }
    }

    return { x: startX, y: startY, w: totalWidth, h: totalHeight, inView: true }
}

/**
 * 获取区域在可见视图中的信息（方法2）
 * @param {Object} area - 区域对象
 * @returns {Object} 包含位置和尺寸信息的对象
 */
export function _getAreaInviewInfo2(area) {

    let { s: { c: startCol, r: startRow }, e: { c: endCol, r: endRow } } = area;

    const viewCol = this.vi.colsIndex
    const viewRow = this.vi.rowsIndex

    // area不在可视区域内就不需要计算了
    if (endRow < viewRow[0] || startRow > viewRow[viewRow.length - 1] || endCol < viewCol[0] || startCol > viewCol[viewCol.length - 1]) return { inView: false }

    // 使用视图信息和滚动位置来确定区域的起始点，xy坐标原点要计算行列坐标的长度
    let totalWidth = 0;
    let totalHeight = 0;

    // 计算区域长度
    for (let r = startRow; r <= endRow; r++) {
        const row = this.getRow(r);
        if (!row.hidden) totalHeight += row.height;
    }
    for (let c = startCol; c <= endCol; c++) {
        const col = this.getCol(c);
        if (!col.hidden) totalWidth += col.width;
    }

    // 计算xy起始坐标
    // 计算起始X坐标
    let offsetX = 0;
    const firstViewCol = viewCol[0];
    if (startCol >= firstViewCol) {
        // 区域左边在可视区内或右侧
        for (let c = firstViewCol; c < startCol; c++) {
            const col = this.getCol(c);
            if (!col.hidden) offsetX += col.width;
        }
    } else {
        // 区域左边在可视区左侧，可能为负数
        for (let c = startCol; c < firstViewCol; c++) {
            const col = this.getCol(c);
            if (!col.hidden) offsetX -= col.width;
        }
    }
    let startX = this.indexWidth + offsetX;

    // 计算起始Y坐标
    let offsetY = 0;
    const firstViewRow = viewRow[0];
    if (startRow >= firstViewRow) {
        for (let r = firstViewRow; r < startRow; r++) {
            const row = this.getRow(r);
            if (!row.hidden) offsetY += row.height;
        }
    } else {
        for (let r = startRow; r < firstViewRow; r++) {
            const row = this.getRow(r);
            if (!row.hidden) offsetY -= row.height;
        }
    }
    let startY = this.headHeight + offsetY;

    return { x: startX, y: startY, w: totalWidth, h: totalHeight, inView: true }
}

/**
 * 更新此表在画布中需要展示的索引信息
 * @param {Number} canvasWidth - 画布宽度
 * @param {Number} canvasHeight - 画布高度
 */
export function _updView(canvasWidth, canvasHeight) {
    // 使用传入的宽高或默认使用画布的宽高（逻辑像素）
    const viewWidth = canvasWidth ?? this.Canvas?.viewWidth ?? this._handleLayer.offsetWidth;
    const viewHeight = canvasHeight ?? this.Canvas?.viewHeight ?? this._handleLayer.offsetHeight;

    // 找到上下左右最后一个非隐藏的行列索引
    let side = { t: null, l: null, r: null, b: null }
    for (let i = this.colCount - 1; i >= 0; i--) {
        if (!this.getCol(i).hidden) {
            side.r = i
            break
        }
    }
    for (let i = this.rowCount - 1; i >= 0; i--) {
        if (!this.getRow(i).hidden) {
            side.b = i
            break
        }
    }
    for (let i = 0; i < this.colCount; i++) {
        if (!this.getCol(i).hidden) {
            side.l = i;
            break;
        }
    }
    for (let i = 0; i < this.rowCount; i++) {
        if (!this.getRow(i).hidden) {
            side.t = i;
            break;
        }
    }
    let vi = this.vi ?? { rowArr: [side.t], colArr: [side.l] };
    // 冻结时限制滚动范围：非冻结区域起始位置不能小于冻结数量
    const minRow = Math.max(side.t, this._frozenRows || 0);
    const minCol = Math.max(side.l, this._frozenCols || 0);
    vi.rowArr[0] = Math.min(Math.max(vi.rowArr[0], minRow), side.b);
    vi.colArr[0] = Math.min(Math.max(vi.colArr[0], minCol), side.r);
    const newVi = {
        frozenCols: [],
        frozenRows: [],
        rowArr: [],
        colArr: [],
        rowsIndex: [],
        colsIndex: [],
        w: this.indexWidth,
        h: this.headHeight,
        fw: 0, // 冻结的宽
        fh: 0, // 冻结的高
        side
    }
    // 处理冻结 - 从 freezeStart 开始渲染到 Split-1
    const xsNum = this.frozenCols
    const ysNum = this.frozenRows
    const freezeStartCol = this._freezeStartCol || 0;
    const freezeStartRow = this._freezeStartRow || 0;

    if (xsNum > 0 && freezeStartCol < xsNum) {
        for (let i = freezeStartCol; i < xsNum; i++) {
            const colInfo = this.getCol(i);
            if (colInfo.hidden) continue;
            newVi.w += colInfo.width;
            newVi.frozenCols.push(i);
        }
    }
    if (ysNum > 0 && freezeStartRow < ysNum) {
        for (let i = freezeStartRow; i < ysNum; i++) {
            const rowInfo = this.getRow(i);
            if (rowInfo.hidden) continue;
            newVi.h += rowInfo.height;
            newVi.frozenRows.push(i);
        }
    }
    newVi.fw = newVi.w
    newVi.fh = newVi.h
    // 获取结束列
    for (let i = vi.colArr[0]; i < this.colCount; i++) {
        const colInfo = this.getCol(i)
        if (colInfo.hidden || i < xsNum) continue
        newVi.w += colInfo.width
        newVi.colArr.push(i) // 此列没隐藏加入计算
        if (newVi.w > viewWidth) break
    }
    // 获取结束行
    for (let i = vi.rowArr[0]; i < this.rowCount; i++) {
        const rowInfo = this.getRow(i);
        if (rowInfo.hidden || i < ysNum) continue
        newVi.h += rowInfo.height
        newVi.rowArr.push(i) // 此行没隐藏加入计算
        if (newVi.h > viewHeight) break
    }
    newVi.rowsIndex = [...newVi.frozenRows, ...newVi.rowArr]
    newVi.colsIndex = [...newVi.frozenCols, ...newVi.colArr]
    this.vi = newVi;
}

/**
 * 往前找指定距离非隐藏的行的索引
 * @param {Number} r - 当前行索引
 * @param {Number} number - 往前查找的距离
 * @returns {Number} 找到的行索引
 */
export function _findBeforeShowRow(r, number) {
    let count = 0
    let res = r
    for (let i = r - 1; i >= 0; i--) {
        if (!this.getRow(i).hidden) {
            count++;
            res = i
        }
        if (count >= number) return i
    }
    return res
}

/**
 * 往前找指定距离非隐藏的列的索引
 * @param {Number} c - 当前列索引
 * @param {Number} number - 往前查找的距离
 * @returns {Number} 找到的列索引
 */
export function _findBeforeShowCol(c, number) {
    let count = 0
    let res = c
    for (let i = c - 1; i >= 0; i--) {
        if (!this.getCol(i).hidden) {
            count++;
            res = i
        }
        if (count >= number) return i
    }
    return res
}

/**
 * 往后找指定距离非隐藏的行的索引
 * @param {Number} r - 当前行索引
 * @param {Number} number - 往后查找的距离
 * @returns {Number} 找到的行索引
 */
export function _findNextShowRow(r, number) {
    let count = 0
    let res = r
    for (let i = r + 1; i < this.rowCount; i++) {
        if (!this.getRow(i).hidden) {
            count++;
            res = i
        }
        if (count >= number) return i
    }
    return res
}

/**
 * 往后找指定距离非隐藏的列的索引
 * @param {Number} c - 当前列索引
 * @param {Number} number - 往后查找的距离
 * @returns {Number} 找到的列索引
 */
export function _findNextShowCol(c, number) {
    let count = 0
    let res = c
    for (let i = c + 1; i < this.colCount; i++) {
        if (!this.getCol(i).hidden) {
            count++;
            res = i
        }
        if (count >= number) return i
    }
    return res
}

/**
 * 获取所有行的总高度（缓存优化）
 * @returns {Number} 总高度
 */
export function getTotalHeight() {
    if (this._totalHeightCache !== undefined) return this._totalHeightCache;
    let total = 0;
    for (let i = 0; i < this.rowCount; i++) {
        if (!this.getRow(i).hidden) total += this.getRow(i).height;
    }
    this._totalHeightCache = total;
    return total;
}

/**
 * 获取所有列的总宽度（缓存优化）
 * @returns {Number} 总宽度
 */
export function getTotalWidth() {
    if (this._totalWidthCache !== undefined) return this._totalWidthCache;
    let total = 0;
    for (let i = 0; i < this.colCount; i++) {
        if (!this.getCol(i).hidden) total += this.getCol(i).width;
    }
    this._totalWidthCache = total;
    return total;
}

/**
 * 清除尺寸缓存（行列变化时调用）
 */
export function _clearSizeCache() {
    this._totalHeightCache = undefined;
    this._totalWidthCache = undefined;
}

/**
 * 获取指定行之前所有行的高度总和
 * @param {Number} rowIndex - 行索引
 * @returns {Number} 高度总和
 */
export function getScrollTop(rowIndex) {
    let total = 0;
    for (let i = 0; i < rowIndex && i < this.rowCount; i++) {
        if (!this.getRow(i).hidden) total += this.getRow(i).height;
    }
    return total;
}

/**
 * 获取指定列之前所有列的宽度总和
 * @param {Number} colIndex - 列索引
 * @returns {Number} 宽度总和
 */
export function getScrollLeft(colIndex) {
    let total = 0;
    for (let i = 0; i < colIndex && i < this.colCount; i++) {
        if (!this.getCol(i).hidden) total += this.getCol(i).width;
    }
    return total;
}

/**
 * 根据垂直滚动像素位置找到对应行索引
 * @param {Number} scrollTop - 滚动像素位置
 * @returns {Number} 行索引
 */
export function getRowIndexByScrollTop(scrollTop) {
    let total = 0;
    for (let i = 0; i < this.rowCount; i++) {
        if (this.getRow(i).hidden) continue;
        total += this.getRow(i).height;
        if (total > scrollTop) return i;
    }
    return Math.max(0, this.rowCount - 1);
}

/**
 * 根据水平滚动像素位置找到对应列索引
 * @param {Number} scrollLeft - 滚动像素位置
 * @returns {Number} 列索引
 */
export function getColIndexByScrollLeft(scrollLeft) {
    let total = 0;
    for (let i = 0; i < this.colCount; i++) {
        if (this.getCol(i).hidden) continue;
        total += this.getCol(i).width;
        if (total > scrollLeft) return i;
    }
    return Math.max(0, this.colCount - 1);
}
