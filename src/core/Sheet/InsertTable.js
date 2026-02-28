/**
 * 插入表模块
 * 负责处理从指定位置插入表格数据的功能
 */

/**
 * 从指定位置开始，插入一个表
 * @param {Array} arr - 表格数据数组
 * @param {Object|String} pos - 插入位置
 * @param {Object} options - 配置选项 {align, border, width, height, background, color}
 */
export function insertTable(arr, pos, options = {}) {
    if (!Array.isArray(arr) || !pos) return;
    if (typeof pos == 'string') pos = this.Utils.cellStrToNum(pos);

    const alignMap = { l: 'left', c: 'center', r: 'right' };

    // 计算表格列数，考虑合并单元格影响
    let maxCol = 0;
    arr.forEach(row => {
        let colIndex = 0;
        let skip = 0; // 用于跳过已合并的列数
        for (let i = 0; i < row.length; i++) {
            if (skip > 0) {
                skip--; // 如果需要跳过列，减一
                continue; // 跳过当前单元格
            }

            const cell = row[i];
            const mergeCols = cell?.mr ?? 0; // 获取水平合并的列数
            colIndex += (mergeCols + 1); // 当前单元格占据的列数

            // 如果当前单元格有合并列，后面的单元格不参与数量计算
            skip = mergeCols;
        }
        maxCol = Math.max(maxCol, colIndex);
    });

    const readyMerge = [];

    for (let rIndex = 0; rIndex < arr.length; rIndex++) {
        let item = arr[rIndex] ?? []
        for (let cIndex = 0; cIndex < maxCol; cIndex++) {

            const r = pos.r + rIndex;
            const c = pos.c + cIndex;

            const cell = this.getCell(r, c);

            const borderStyle = { style: 'thin', color: '#000000' };
            if (options.border) cell.border = { top: borderStyle, left: borderStyle, bottom: borderStyle, right: borderStyle };
            if (options.align) cell.alignment = { vertical: 'center', horizontal: options.align };

            const tCell = item[cIndex] ?? {};
            if (Object.prototype.toString.call(tCell) != '[object Object]') {
                cell.editVal = tCell;
                continue;
            }

            cell.editVal = tCell.v;
            if (tCell.fg) cell.fill = { fgColor: tCell.fg };
            if (tCell.c) cell.font = { color: tCell.c };

            if (tCell.a) cell.alignment = { horizontal: alignMap[tCell.a] };
            if (tCell.b) cell.font = { bold: true };
            if (tCell.s) cell.font = { size: Number(tCell.s) }

            if (tCell.mr || tCell.mb) {
                readyMerge.push({
                    s: { r, c },
                    e: { r: r + (tCell.mb ?? 0), c: c + (tCell.mr ?? 0) }
                });
            }
        }

        if (item?.[0]?.h || options.height) this.getRow(pos.r + rIndex).height = item?.[0]?.h ?? options.height;
    }

    // 应用全局列宽设置
    for (let index = 0; index < maxCol; index++) {
        const item = arr[0][index] ?? {};
        const w = item.w ?? options.width; // 如果单元格未定义宽度，使用默认宽度
        if (w) this.getCol(pos.c + index).width = w;
    }

    readyMerge.forEach(m => this.mergeCells(m));

    const range = { s: pos, e: { r: pos.r + arr.length - 1, c: pos.c + maxCol - 1 } };
    return range;
}
