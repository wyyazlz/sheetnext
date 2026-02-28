/**
 * 剪贴板核心模块 - Excel 一比一复刻
 * 支持：复制/剪切/粘贴、选择性粘贴、跨 sheet 操作
 * 高性能设计：批量操作、最小化重绘
 */

import { PasteMode, PasteOperation } from '../../enum/index.js';
export { PasteMode, PasteOperation };

// 内部剪贴板数据存储（跨 sheet 共享）
let clipboardData = null;

function emitIfListeners(SN, event, data) {
    if (!SN?.hasListeners?.(event)) return null;
    return SN.Event.emit(event, data);
}

/**
 * 复制区域数据
 * @param {Object} area - 复制区域 { s: {r, c}, e: {r, c} }
 * @param {boolean} isCut - 是否剪切
 */
export function copy(area = null, isCut = false) {
    const sheet = this;
    if (!area) {
        area = sheet.activeAreas[sheet.activeAreas.length - 1];
    }
    if (typeof area === 'string') {
        area = sheet.rangeStrToNum(area);
    }

    const eventBase = isCut ? 'Cut' : 'Copy';
    const payload = {
        sheet,
        area: JSON.parse(JSON.stringify(area)),
        isCut
    };
    const beforeEvent = emitIfListeners(sheet.SN, `before${eventBase}`, payload);
    if (beforeEvent?.canceled) return null;

    // 收集单元格数据
    const data = {
        sheetName: sheet.name,
        area: JSON.parse(JSON.stringify(area)),
        isCut,
        cells: [],      // 单元格数据二维数组
        merges: [],     // 合并单元格信息
        colWidths: [],  // 列宽
        rowHeights: [], // 行高
        comments: [],   // 批注
        timestamp: Date.now()
    };

    const rowCount = area.e.r - area.s.r + 1;
    const colCount = area.e.c - area.s.c + 1;

    // 收集列宽
    for (let c = area.s.c; c <= area.e.c; c++) {
        data.colWidths.push(sheet.getCol(c).width);
    }

    // 收集行高
    for (let r = area.s.r; r <= area.e.r; r++) {
        data.rowHeights.push(sheet.getRow(r).height);
    }

    // 收集单元格数据
    for (let r = area.s.r; r <= area.e.r; r++) {
        const rowData = [];
        for (let c = area.s.c; c <= area.e.c; c++) {
            const cell = sheet.getCell(r, c);
            rowData.push(copyCellData(cell, sheet));
        }
        data.cells.push(rowData);
    }

    // 收集合并单元格信息（相对坐标）
    sheet.merges.forEach(merge => {
        if (merge.s.r >= area.s.r && merge.e.r <= area.e.r &&
            merge.s.c >= area.s.c && merge.e.c <= area.e.c) {
            data.merges.push({
                s: { r: merge.s.r - area.s.r, c: merge.s.c - area.s.c },
                e: { r: merge.e.r - area.s.r, c: merge.e.c - area.s.c }
            });
        }
    });

    // 收集批注（相对坐标）
    if (sheet.Comment?.list) {
        sheet.Comment.list.forEach(comment => {
            const pos = comment.cellPos;
            if (pos.r >= area.s.r && pos.r <= area.e.r &&
                pos.c >= area.s.c && pos.c <= area.e.c) {
                data.comments.push({
                    relR: pos.r - area.s.r,
                    relC: pos.c - area.s.c,
                    author: comment.author,
                    text: comment.text,
                    width: comment.width,
                    height: comment.height
                });
            }
        });
    }

    // 存储到内部剪贴板
    clipboardData = data;

    emitIfListeners(sheet.SN, `after${eventBase}`, { ...payload, data });
    return data;
}

/**
 * 剪切区域数据
 * @param {Object} area - 剪切区域
 */
export function cut(area = null) {
    return this.copy(area, true);
}

/**
 * 粘贴数据到目标区域
 * @param {Object} targetArea - 目标区域起始位置 { r, c } 或完整区域
 * @param {Object} options - 粘贴选项
 * @param {string} options.mode - 粘贴模式 (PasteMode)
 * @param {string} options.operation - 运算模式 (PasteOperation)
 * @param {boolean} options.skipBlanks - 跳过空单元格
 * @param {boolean} options.transpose - 转置
 * @param {Object} options.externalData - 外部剪贴板数据（从系统剪贴板解析）
 */
export function paste(targetArea = null, options = {}) {
    const sheet = this;
    const {
        mode = PasteMode.ALL,
        operation = PasteOperation.NONE,
        skipBlanks = false,
        transpose = false,
        externalData = null
    } = options;

    // 确定目标位置
    if (!targetArea) {
        targetArea = { r: sheet.activeCell.r, c: sheet.activeCell.c };
    }
    const targetR = targetArea.r ?? targetArea.s?.r ?? 0;
    const targetC = targetArea.c ?? targetArea.s?.c ?? 0;

    // 获取数据源
    const data = externalData || clipboardData;
    if (!data || !data.cells || data.cells.length === 0) {
        return false;
    }

    // 计算实际尺寸（考虑转置）
    const actualTranspose = transpose || mode === PasteMode.TRANSPOSE;
    const srcRowCount = data.cells.length;
    const srcColCount = data.cells[0]?.length || 0;
    const rowCount = actualTranspose ? srcColCount : srcRowCount;
    const colCount = actualTranspose ? srcRowCount : srcColCount;

    // 计算目标区域
    const pasteArea = {
        s: { r: targetR, c: targetC },
        e: { r: targetR + rowCount - 1, c: targetC + colCount - 1 }
    };
    const payload = {
        sheet,
        targetArea: pasteArea,
        options: { mode, operation, skipBlanks, transpose: actualTranspose },
        data
    };
    const beforeEvent = emitIfListeners(sheet.SN, 'beforePaste', payload);
    if (beforeEvent?.canceled) return false;

    const valueModes = new Set([
        PasteMode.ALL,
        PasteMode.ALL_MERGE_CONDITIONAL,
        PasteMode.VALUE,
        PasteMode.FORMULA,
        PasteMode.NO_BORDER,
        PasteMode.TRANSPOSE,
        PasteMode.VALUE_NUMBER_FORMAT,
        PasteMode.FORMULA_NUMBER_FORMAT
    ]);
    const formatModes = new Set([
        PasteMode.ALL,
        PasteMode.ALL_MERGE_CONDITIONAL,
        PasteMode.FORMAT,
        PasteMode.NO_BORDER,
        PasteMode.TRANSPOSE,
        PasteMode.VALUE_NUMBER_FORMAT,
        PasteMode.FORMULA_NUMBER_FORMAT,
        PasteMode.VALIDATION
    ]);
    const needsEdit = valueModes.has(mode) || operation !== PasteOperation.NONE;
    const needsFormat = formatModes.has(mode) || shouldPasteMerges(mode);
    const needsColumnFormat = mode === PasteMode.ALL || mode === PasteMode.COLUMN_WIDTH;
    const needsComments = shouldPasteComments(mode);
    const needsHyperlink = [
        PasteMode.ALL,
        PasteMode.ALL_MERGE_CONDITIONAL,
        PasteMode.NO_BORDER,
        PasteMode.TRANSPOSE
    ].includes(mode);

    // 检查是否与源区域重叠（剪切时）
    if (data.isCut && data.sheetName === sheet.name) {
        const srcArea = data.area;
        if (areasOverlap(srcArea, pasteArea)) {
            sheet.SN.Utils.toast(sheet.SN.t('core.sheet.clipboard.toast.msg001'));
            return false;
        }
    }

    // 预处理：清除目标区域的合并单元格
    clearTargetMerges(sheet, pasteArea);

    // 批量粘贴单元格数据
    for (let r = 0; r < rowCount; r++) {
        for (let c = 0; c < colCount; c++) {
            // 获取源数据（考虑转置）
            const srcR = actualTranspose ? c : r;
            const srcC = actualTranspose ? r : c;
            const srcCell = data.cells[srcR]?.[srcC];

            if (!srcCell) continue;

            // 跳过空单元格
            if (skipBlanks && isEmpty(srcCell)) continue;

            const targetCell = sheet.getCell(targetR + r, targetC + c);

            // 根据模式粘贴
            pasteCellByMode(targetCell, srcCell, mode, operation, sheet);
        }
    }

    // 粘贴合并单元格（非仅值/仅公式模式）
    if (shouldPasteMerges(mode) && data.merges.length > 0) {
        data.merges.forEach(merge => {
            let newMerge;
            if (actualTranspose) {
                newMerge = {
                    s: { r: targetR + merge.s.c, c: targetC + merge.s.r },
                    e: { r: targetR + merge.e.c, c: targetC + merge.e.r }
                };
            } else {
                newMerge = {
                    s: { r: targetR + merge.s.r, c: targetC + merge.s.c },
                    e: { r: targetR + merge.e.r, c: targetC + merge.e.c }
                };
            }
            try {
                sheet.mergeCells(newMerge);
            } catch (e) {
                // 静默忽略合并失败
            }
        });
    }

    // 粘贴列宽
    if (mode === PasteMode.ALL || mode === PasteMode.COLUMN_WIDTH) {
        if (actualTranspose) {
            // 转置时行高变列宽
            data.rowHeights.forEach((h, i) => {
                if (h) sheet.getCol(targetC + i).width = h;
            });
        } else {
            data.colWidths.forEach((w, i) => {
                if (w) sheet.getCol(targetC + i).width = w;
            });
        }
    }

    // 粘贴批注
    if (shouldPasteComments(mode) && data.comments.length > 0) {
        data.comments.forEach(comment => {
            let newR, newC;
            if (actualTranspose) {
                newR = targetR + comment.relC;
                newC = targetC + comment.relR;
            } else {
                newR = targetR + comment.relR;
                newC = targetC + comment.relC;
            }
            const cellRef = sheet.SN.Utils.numToCell(newR, newC);
            sheet.Comment.add({
                cellRef,
                author: comment.author,
                text: comment.text,
                width: comment.width,
                height: comment.height
            });
        });
    }

    // 如果是剪切，清除源区域
    if (data.isCut && !externalData) {
        const sourceSheet = sheet.SN.getSheet(data.sheetName);
        if (sourceSheet) {
            const cleared = clearArea(sourceSheet, data.area);
            if (!cleared) return true;
        }
        // 清除内部剪贴板（剪切只能粘贴一次）
        clipboardData = null;
    }

    emitIfListeners(sheet.SN, 'afterPaste', { ...payload, success: true });
    return true;
}

/**
 * 清除剪贴板
 */
export function clearClipboard() {
    const sheet = this;
    const beforeEvent = emitIfListeners(sheet.SN, 'beforeClearClipboard', { sheet });
    if (beforeEvent?.canceled) return;
    clipboardData = null;
    emitIfListeners(sheet.SN, 'afterClearClipboard', { sheet });
}

/**
 * 获取当前剪贴板数据
 */
export function getClipboardData() {
    return clipboardData;
}

/**
 * 检查是否有剪贴板数据
 */
export function hasClipboardData() {
    return clipboardData !== null;
}

// ==================== 内部辅助函数 ====================

/**
 * 复制单元格数据
 */
function copyCellData(cell, sheet) {
    return {
        editVal: cell._editVal,
        calcVal: cell._calcVal,
        type: cell._type,
        style: cell._style ? JSON.parse(JSON.stringify(cell._style)) : null,
        hyperlink: cell._hyperlink ? JSON.parse(JSON.stringify(cell._hyperlink)) : null,
        dataValidation: cell._dataValidation ? JSON.parse(JSON.stringify(cell._dataValidation)) : null,
        isMerged: cell.isMerged,
        master: cell.master ? { r: cell.master.r, c: cell.master.c } : null
    };
}

/**
 * 根据模式粘贴单元格
 */
function pasteCellByMode(targetCell, srcCell, mode, operation, sheet) {
    // 先计算运算结果（如果需要）
    let valueToSet = srcCell.editVal;
    if (operation !== PasteOperation.NONE && typeof srcCell.calcVal === 'number') {
        const targetVal = typeof targetCell.calcVal === 'number' ? targetCell.calcVal : 0;
        const srcVal = srcCell.calcVal;
        switch (operation) {
            case PasteOperation.ADD:
                valueToSet = targetVal + srcVal;
                break;
            case PasteOperation.SUBTRACT:
                valueToSet = targetVal - srcVal;
                break;
            case PasteOperation.MULTIPLY:
                valueToSet = targetVal * srcVal;
                break;
            case PasteOperation.DIVIDE:
                valueToSet = srcVal !== 0 ? targetVal / srcVal : '#DIV/0!';
                break;
        }
    }

    switch (mode) {
        case PasteMode.ALL:
        case PasteMode.ALL_MERGE_CONDITIONAL:
            // 全部粘贴
            setValueWithoutUndo(targetCell, valueToSet);
            if (srcCell.style) targetCell._style = JSON.parse(JSON.stringify(srcCell.style));
            if (srcCell.hyperlink) targetCell._hyperlink = JSON.parse(JSON.stringify(srcCell.hyperlink));
            if (srcCell.dataValidation) targetCell._dataValidation = JSON.parse(JSON.stringify(srcCell.dataValidation));
            break;

        case PasteMode.VALUE:
            // 仅值（公式转为计算结果）
            const val = srcCell.calcVal !== undefined ? srcCell.calcVal : srcCell.editVal;
            setValueWithoutUndo(targetCell, operation !== PasteOperation.NONE ? valueToSet : val);
            break;

        case PasteMode.FORMULA:
            // 仅公式（需要调整公式中的引用）
            if (typeof srcCell.editVal === 'string' && srcCell.editVal.startsWith('=')) {
                // TODO: 调整公式引用
                setValueWithoutUndo(targetCell, srcCell.editVal);
            } else {
                setValueWithoutUndo(targetCell, srcCell.editVal);
            }
            break;

        case PasteMode.FORMAT:
            // 仅格式
            if (srcCell.style) targetCell._style = JSON.parse(JSON.stringify(srcCell.style));
            break;

        case PasteMode.NO_BORDER:
            // 无边框
            setValueWithoutUndo(targetCell, valueToSet);
            if (srcCell.style) {
                const styleWithoutBorder = JSON.parse(JSON.stringify(srcCell.style));
                delete styleWithoutBorder.border;
                targetCell._style = styleWithoutBorder;
            }
            if (srcCell.hyperlink) targetCell._hyperlink = JSON.parse(JSON.stringify(srcCell.hyperlink));
            if (srcCell.dataValidation) targetCell._dataValidation = JSON.parse(JSON.stringify(srcCell.dataValidation));
            break;

        case PasteMode.VALUE_NUMBER_FORMAT:
            // 值和数字格式
            const valNum = srcCell.calcVal !== undefined ? srcCell.calcVal : srcCell.editVal;
            setValueWithoutUndo(targetCell, valNum);
            if (srcCell.style?.numFmt) {
                if (!targetCell._style) targetCell._style = {};
                targetCell._style.numFmt = srcCell.style.numFmt;
            }
            break;

        case PasteMode.FORMULA_NUMBER_FORMAT:
            // 公式和数字格式
            setValueWithoutUndo(targetCell, srcCell.editVal);
            if (srcCell.style?.numFmt) {
                if (!targetCell._style) targetCell._style = {};
                targetCell._style.numFmt = srcCell.style.numFmt;
            }
            break;

        case PasteMode.COLUMN_WIDTH:
            // 列宽（在外层处理）
            break;

        case PasteMode.Comment:
            // 批注（在外层处理）
            break;

        case PasteMode.VALIDATION:
            // 仅数据验证
            if (srcCell.dataValidation) {
                targetCell._dataValidation = JSON.parse(JSON.stringify(srcCell.dataValidation));
            }
            break;

        case PasteMode.TRANSPOSE:
            // 转置（在外层处理坐标转换，这里按 ALL 处理）
            setValueWithoutUndo(targetCell, valueToSet);
            if (srcCell.style) targetCell._style = JSON.parse(JSON.stringify(srcCell.style));
            if (srcCell.hyperlink) targetCell._hyperlink = JSON.parse(JSON.stringify(srcCell.hyperlink));
            if (srcCell.dataValidation) targetCell._dataValidation = JSON.parse(JSON.stringify(srcCell.dataValidation));
            break;
    }
}

/**
 * 设置值（不触发撤销记录，批量操作结束后统一记录）
 */
function setValueWithoutUndo(cell, val) {
    cell._calcVal = undefined;
    cell._showVal = undefined;

    // 空值处理
    if (val == null || val === '' || val === undefined || Number.isNaN(val)) {
        cell._type = undefined;
        cell._editVal = "";
    }
    // 已经是数字类型
    else if (typeof val === 'number') {
        cell._editVal = val;
        cell._type = undefined;
    }
    // 公式
    else if (typeof val === 'string' && val.trim().startsWith('=')) {
        cell._editVal = val.trim();
        cell._type = undefined;
    }
    // 布尔值
    else if (typeof val === 'boolean') {
        cell._editVal = val;
        cell._type = 'boolean';
    }
    // 字符串数字（需要严格判断）
    else if (typeof val === 'string' && val.trim() !== '' && !isNaN(Number(val))) {
        cell._editVal = parseFloat(val);
        cell._type = undefined;
    }
    // 其他字符串
    else {
        cell._editVal = val;
        cell._type = 'string';
    }
}

/**
 * 判断是否应粘贴合并单元格
 */
function shouldPasteMerges(mode) {
    return [
        PasteMode.ALL,
        PasteMode.ALL_MERGE_CONDITIONAL,
        PasteMode.FORMAT,
        PasteMode.NO_BORDER,
        PasteMode.TRANSPOSE
    ].includes(mode);
}

/**
 * 判断是否应粘贴批注
 */
function shouldPasteComments(mode) {
    return [
        PasteMode.ALL,
        PasteMode.ALL_MERGE_CONDITIONAL,
        PasteMode.Comment
    ].includes(mode);
}

/**
 * 清除目标区域的合并单元格
 */
function clearTargetMerges(sheet, area) {
    const toRemove = [];
    sheet.merges.forEach((merge, index) => {
        // 检查是否与目标区域有交集
        if (!(merge.e.r < area.s.r || merge.s.r > area.e.r ||
              merge.e.c < area.s.c || merge.s.c > area.e.c)) {
            toRemove.push(merge);
        }
    });
    toRemove.forEach(merge => {
        sheet.unMergeCells({ r: merge.s.r, c: merge.s.c });
    });
}

/**
 * 清除区域数据（剪切后清除源区域）
 */
function clearArea(sheet, area) {
    sheet.eachCells(area, (r, c) => {
        const cell = sheet.getCell(r, c);
        cell._editVal = "";
        cell._calcVal = undefined;
        cell._showVal = undefined;
        cell._style = undefined;
        cell._hyperlink = null;
        cell._dataValidation = null;
    });

    // 清除合并单元格
    const toRemove = [];
    sheet.merges.forEach(merge => {
        if (merge.s.r >= area.s.r && merge.e.r <= area.e.r &&
            merge.s.c >= area.s.c && merge.e.c <= area.e.c) {
            toRemove.push(merge);
        }
    });
    toRemove.forEach(merge => {
        sheet.unMergeCells({ r: merge.s.r, c: merge.s.c });
    });

    // 清除批注
    if (sheet.Comment?.list) {
        const commentsToRemove = [];
        sheet.Comment.list.forEach(comment => {
            const pos = comment.cellPos;
            if (pos.r >= area.s.r && pos.r <= area.e.r &&
                pos.c >= area.s.c && pos.c <= area.e.c) {
                commentsToRemove.push(comment.cellRef);
            }
        });
        commentsToRemove.forEach(ref => {
            sheet.Comment.remove(ref);
        });
    }

    return true;
}

/**
 * 判断两个区域是否重叠
 */
function areasOverlap(area1, area2) {
    return !(area1.e.r < area2.s.r || area1.s.r > area2.e.r ||
             area1.e.c < area2.s.c || area1.s.c > area2.e.c);
}

/**
 * 判断单元格是否为空
 */
function isEmpty(cellData) {
    return cellData.editVal === "" || cellData.editVal === null || cellData.editVal === undefined;
}

/**
 * 获取粘贴模式标签
 */
function getPasteModeLabel(mode) {
    const labels = {
        [PasteMode.ALL]: '全部',
        [PasteMode.VALUE]: '值',
        [PasteMode.FORMULA]: '公式',
        [PasteMode.FORMAT]: '格式',
        [PasteMode.NO_BORDER]: '无边框',
        [PasteMode.COLUMN_WIDTH]: '列宽',
        [PasteMode.Comment]: '批注',
        [PasteMode.VALIDATION]: '数据验证',
        [PasteMode.TRANSPOSE]: '转置',
        [PasteMode.VALUE_NUMBER_FORMAT]: '值和数字格式',
        [PasteMode.ALL_MERGE_CONDITIONAL]: '所有使用源主题',
        [PasteMode.FORMULA_NUMBER_FORMAT]: '公式和数字格式'
    };
    return labels[mode] || mode;
}

/**
 * 获取运算模式标签
 */
function getOperationLabel(op) {
    const labels = {
        [PasteOperation.NONE]: '无',
        [PasteOperation.ADD]: '加',
        [PasteOperation.SUBTRACT]: '减',
        [PasteOperation.MULTIPLY]: '乘',
        [PasteOperation.DIVIDE]: '除'
    };
    return labels[op] || op;
}

/**
 * 从系统剪贴板解析数据（用于外部粘贴）
 * @param {Object} parsed - parseHTMLStyle 解析的结果
 * @returns {Object} 内部剪贴板格式数据
 */
export function fromParsedHTML(parsed) {
    if (!parsed || !parsed.data) return null;

    const cells = parsed.data.map((row, rIndex) => {
        return row.map((value, cIndex) => {
            // 严格判断是否为数字
            const isNumeric = value !== '' && value != null && !isNaN(Number(value));
            return {
                editVal: value,
                calcVal: isNumeric ? Number(value) : value,
                type: isNumeric ? 'number' : 'string',
                style: parsed.styles?.[rIndex]?.[cIndex] || null,
                hyperlink: null,
                dataValidation: null,
                isMerged: false,
                master: null
            };
        });
    });

    return {
        sheetName: null, // 外部数据
        area: {
            s: { r: 0, c: 0 },
            e: { r: cells.length - 1, c: cells[0]?.length - 1 || 0 }
        },
        isCut: false,
        cells,
        merges: parsed.merges || [],
        colWidths: [],
        rowHeights: [],
        comments: [],
        timestamp: Date.now(),
        isExternal: true
    };
}

/**
 * 从纯文本解析数据
 * @param {string} text - Tab 分隔的文本
 * @returns {Object} 内部剪贴板格式数据
 */
export function fromPlainText(text) {
    if (!text) return null;

    const rows = text.split(/\r\n|\n/).filter(row => row);
    const cells = rows.map(row => {
        return row.split('\t').map(value => {
            // 严格判断是否为数字（排除空字符串、超长数字字符串）
            const isNumeric = value !== '' && value.length <= 15 && !isNaN(Number(value));
            return {
                editVal: isNumeric ? Number(value) : value,
                calcVal: isNumeric ? Number(value) : value,
                type: isNumeric ? 'number' : 'string',
                style: null,
                hyperlink: null,
                dataValidation: null,
                isMerged: false,
                master: null
            };
        });
    });

    return {
        sheetName: null,
        area: {
            s: { r: 0, c: 0 },
            e: { r: cells.length - 1, c: cells[0]?.length - 1 || 0 }
        },
        isCut: false,
        cells,
        merges: [],
        colWidths: [],
        rowHeights: [],
        comments: [],
        timestamp: Date.now(),
        isExternal: true
    };
}
