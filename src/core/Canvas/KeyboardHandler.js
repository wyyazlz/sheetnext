import { isCheckboxControl } from '../Cell/cellControlXml.js';

const ARROW_DIRECTIONS = {
    ArrowUp: 'up',
    ArrowDown: 'down',
    ArrowLeft: 'left',
    ArrowRight: 'right'
};

const MODIFIER_KEYS = new Set([
    'Alt', 'AltGraph', 'CapsLock', 'Control', 'Fn', 'FnLock', 'Hyper',
    'Meta', 'NumLock', 'ScrollLock', 'Shift', 'Super', 'Symbol', 'SymbolLock'
]);

/**
 * Keyboard Event Processing Module
 * Calling the Action Layer's Clipboard Method
 */

// 键盘按下
function insertLineBreakAtCursor(input) {
    const selection = window.getSelection();
    if (!selection) return false;

    let range = null;
    if (selection.rangeCount > 0) {
        const currentRange = selection.getRangeAt(0);
        if (input.contains(currentRange.commonAncestorContainer)) {
            range = currentRange;
        }
    }

    if (!range) {
        range = document.createRange();
        range.selectNodeContents(input);
        range.collapse(false);
        selection.removeAllRanges();
        selection.addRange(range);
    }

    range.deleteContents();
    const lineBreak = document.createElement('br');
    range.insertNode(lineBreak);

    // 末尾的 <br> 在 contenteditable 中不可见，需要追加占位 <br>
    // nextSibling 可能是空文本节点，也需要处理
    const next = lineBreak.nextSibling;
    if (!next || (next.nodeType === 3 && !next.textContent)) {
        input.appendChild(document.createElement('br'));
    }

    const caretRange = document.createRange();
    caretRange.setStartAfter(lineBreak);
    caretRange.collapse(true);
    selection.removeAllRanges();
    selection.addRange(caretRange);
    return true;
}

function normalizeArea(area) {
    if (!area?.s || !area?.e) return null;
    return {
        s: {
            r: Math.min(area.s.r, area.e.r),
            c: Math.min(area.s.c, area.e.c)
        },
        e: {
            r: Math.max(area.s.r, area.e.r),
            c: Math.max(area.s.c, area.e.c)
        }
    };
}

function sameCell(a, b) {
    return a?.r === b?.r && a?.c === b?.c;
}

function sameArea(a, b) {
    return !!a && !!b && sameCell(a.s, b.s) && sameCell(a.e, b.e);
}

function areaContains(area, cell) {
    return !!area && !!cell &&
        cell.r >= area.s.r && cell.r <= area.e.r &&
        cell.c >= area.s.c && cell.c <= area.e.c;
}

function getCellSpan(sheet, cell) {
    const target = sheet.getCell(cell.r, cell.c);
    const master = target.master || cell;
    const merge = target.isMerged || target.master
        ? sheet.merges.find(item => item.s.r === master.r && item.s.c === master.c)
        : null;
    return merge || { s: { ...master }, e: { ...master } };
}

function normalizeTabCell(sheet, cell) {
    const target = sheet.getCell(cell.r, cell.c);
    return target.master ? { ...target.master } : { ...cell };
}

function isTabCellVisible(sheet, cell) {
    const row = sheet.rows[cell.r];
    const pendingRow = !row ? sheet._getPendingXlsxRowState?.(cell.r) : null;
    if (row?.hidden || pendingRow?.hidden || pendingRow?.filterHidden) return false;
    return !sheet.cols[cell.c]?.hidden;
}

function isTabCellSelectable(sheet, cell) {
    return sheet.protection?.isCellSelectable?.(sheet.getCell(cell.r, cell.c)) !== false;
}

function stepInsideArea(sheet, area, cell, reverse) {
    const span = getCellSpan(sheet, cell);
    let r = span.s.r;
    let c = reverse ? span.s.c - 1 : span.e.c + 1;

    if (reverse && c < area.s.c) {
        c = area.e.c;
        r--;
        if (r < area.s.r) r = area.e.r;
    } else if (!reverse && c > area.e.c) {
        c = area.s.c;
        r++;
        if (r > area.e.r) r = area.s.r;
    }

    return { r, c };
}

function findTabTargetInArea(sheet, area, activeCell, reverse) {
    if (!areaContains(area, activeCell)) return null;
    if (sheet.protection?.enabled &&
        !sheet.protection.selectLockedCells &&
        !sheet.protection.selectUnlockedCells) {
        return null;
    }

    const cellCount = (area.e.r - area.s.r + 1) * (area.e.c - area.s.c + 1);
    let cursor = { ...activeCell };
    for (let i = 0; i < cellCount; i++) {
        cursor = stepInsideArea(sheet, area, cursor, reverse);
        const target = normalizeTabCell(sheet, cursor);
        if (sameCell(target, activeCell)) continue;
        if (!areaContains(area, target)) continue;
        if (!isTabCellVisible(sheet, target)) continue;
        if (!isTabCellSelectable(sheet, target)) continue;
        return target;
    }
    return null;
}

function isRestrictedSelection(sheet) {
    return !!sheet.protection?.enabled &&
        (!sheet.protection.selectLockedCells || !sheet.protection.selectUnlockedCells);
}

function findNormalTabTarget(sheet, activeCell, reverse) {
    const span = getCellSpan(sheet, activeCell);
    const boundary = reverse ? span.s.c : span.e.c;
    const c = reverse
        ? sheet._findBeforeShowCol(boundary, 1)
        : sheet._findNextShowCol(boundary, 1);
    const target = normalizeTabCell(sheet, { r: span.s.r, c });

    if (!sameCell(target, activeCell) && isTabCellSelectable(sheet, target)) return target;
    if (!isRestrictedSelection(sheet)) return null;

    const sheetArea = {
        s: { r: 0, c: 0 },
        e: { r: Math.max(0, sheet.rowCount - 1), c: Math.max(0, sheet.colCount - 1) }
    };
    return findTabTargetInArea(sheet, sheetArea, activeCell, reverse);
}

function getTableAppendTarget(sheet, activeCell, reverse, cyclesSelection) {
    if (reverse || cyclesSelection) return null;
    const table = sheet.Table?.getTableAt?.(activeCell.r, activeCell.c);
    if (!table?.range || table.showTotalsRow || !table.isDataCell(activeCell.r, activeCell.c)) return null;
    if (activeCell.r !== table.dataEndRow || activeCell.c !== table.range.e.c) return null;
    if (!sheet.Table.appendRow(table)) return null;
    return { r: table.dataEndRow, c: table.range.s.c };
}

function ensureCellVisible(canvas, sheet, target) {
    const info = sheet.getCellInViewInfo(target.r, target.c);
    if (info.inView) return;

    const vi = sheet.vi || {};
    const bodyLeft = vi.fw ?? sheet.indexWidth;
    const bodyTop = vi.fh ?? sheet.headHeight;
    const bodyWidth = Math.max(1, canvas.viewWidth - bodyLeft);
    const bodyHeight = Math.max(1, canvas.viewHeight - bodyTop);
    let scrollLeft = vi.scrollLeft ?? 0;
    let scrollTop = vi.scrollTop ?? 0;

    if (info.x < bodyLeft) {
        scrollLeft = sheet.getScrollLeft(target.c);
    } else if (info.x + info.w > canvas.viewWidth) {
        scrollLeft = sheet.getScrollLeft(target.c + 1) - bodyWidth;
    }
    if (info.y < bodyTop) {
        scrollTop = sheet.getScrollTop(target.r);
    } else if (info.y + info.h > canvas.viewHeight) {
        scrollTop = sheet.getScrollTop(target.r + 1) - bodyHeight;
    }

    canvas.scrollToPixels(scrollLeft, scrollTop);
}

/**
 * Commit-compatible Excel Tab navigation shared by the cell editor and formula bar.
 * @param {boolean} [reverse=false]
 * @returns {boolean}
 */
export function moveByTab(reverse = false) {
    const sheet = this.activeSheet;
    if (!sheet) return false;

    const activeCell = { ...sheet.activeCell };
    const areas = sheet.activeAreas;
    const area = normalizeArea(areas[areas.length - 1]);
    const activeSpan = getCellSpan(sheet, activeCell);
    const cyclesSelection = !!area &&
        (area.s.r !== area.e.r || area.s.c !== area.e.c) &&
        !sameArea(area, activeSpan);

    const target = getTableAppendTarget(sheet, activeCell, reverse, cyclesSelection)
        || (cyclesSelection
            ? findTabTargetInArea(sheet, area, activeCell, reverse)
            : findNormalTabTarget(sheet, activeCell, reverse));
    if (!target || sameCell(target, activeCell)) return false;

    sheet.activeCell = target;
    if (!sameCell(sheet.activeCell, target)) return false;
    if (!cyclesSelection) sheet.activeAreas = [];
    ensureCellVisible(this, sheet, target);
    this.input?.focus();
    this.r('s');
    return true;
}

export function docKeyDown(event) {
    // 公式编辑态优先处理（联想下拉导航、F4、方向键 Point 引用等）
    if (this.formulaEditor?.handleKeyDown(event)) return;
    if (event.altKey) {
        const isFormulaBarFocused = document.activeElement === this.formulaBar;
        if (event.key === 'Enter' && !isFormulaBarFocused) {
            if (!this.inputEditing && document.activeElement === this.input) {
                event.preventDefault();
                this.showCellInput(false, true);
            }
            if (this.inputEditing) {
                event.preventDefault();
                if (document.activeElement !== this.input) {
                    this.input.focus();
                }
                if (insertLineBreakAtCursor(this.input)) {
                    this.formulaBar.value = this.input.innerText;
                }
            }
        }
        return;
    }
    const sheet = this.activeSheet;
    const area = sheet.activeAreas[sheet.activeAreas.length - 1];

    if (event.key === 'Tab' && document.activeElement === this.input) {
        event.preventDefault();
        if (this.inputEditing) this.updInputValue();
        moveByTab.call(this, event.shiftKey);
        return;
    }

    // F9 - 全局重算（包括 Volatile 函数）
    if (event.key === 'F9') {
        event.preventDefault();
        this.SN.recalculate(event.shiftKey); // Shift+F9 仅重算当前表
        return;
    }

    const selectionDirection = ARROW_DIRECTIONS[event.key];
    if (selectionDirection && event.shiftKey && !this.inputEditing &&
        document.activeElement !== this.formulaBar &&
        document.activeElement === this.input) {
        event.preventDefault();
        const target = sheet.moveSelection(selectionDirection, {
            extend: true,
            jump: event.ctrlKey || event.metaKey
        });
        if (target) {
            ensureCellVisible(this, sheet, target);
            this.input?.focus();
            this.r('s');
        }
        return;
    }

    if (event.ctrlKey) {
        // 快捷输入选中区域（AI 输入框）
        if (document.activeElement.classList.contains('sn-prompt-input') && event.key == 'i') {
            event.preventDefault();
            const inputElement = document.activeElement;
            const start = inputElement.selectionStart;
            const end = inputElement.selectionEnd;
            const text = inputElement.value;
            const areaText = this.areaInput?.value || '';
            const newText = text.slice(0, start) + areaText + text.slice(end);
            inputElement.value = newText;
            inputElement.selectionStart = inputElement.selectionEnd = start + areaText.length;
            return;
        }

        // 检查焦点状态
        if (document.activeElement !== this.input ||
            this.inputEditing ||
            window.getSelection().toString()) {
            return;
        }

        // Ctrl+C 复制
        if (event.key == 'c') {
            event.preventDefault();
            this.SN.Action.copy();
        }
        // Ctrl+X 剪切
        else if (event.key == 'x') {
            event.preventDefault();
            this.SN.Action.cut();
            this.r(); // 刷新显示剪切状态
        }
        // Ctrl+V 粘贴
        else if (event.key == 'v') {
            // 如果有内部剪贴板数据，使用内部粘贴（支持完整功能）
            if (sheet.hasClipboardData()) {
                event.preventDefault();
                this.SN.Action.paste();
            }
            // 否则让浏览器触发 paste 事件，由 docPaste 处理
        }
        // Ctrl+Z 撤销
        else if (event.key === 'z') {
            this.SN.UndoRedo.undo();
            event.preventDefault();
        }
        // Ctrl+Y 重做
        else if (event.key === 'y') {
            this.SN.UndoRedo.redo();
            event.preventDefault();
        }
        // Ctrl+Shift+V 选择性粘贴
        else if (event.key === 'V' && event.shiftKey) {
            event.preventDefault();
            // 触发选择性粘贴对话框
            this.SN.Action.pasteSpecial();
        }

        return;
    }

    // 上下左右移动活动的单元格
    if (['ArrowUp', 'ArrowRight', 'ArrowLeft', 'ArrowDown', 'Enter'].includes(event.key) &&
        document.activeElement !== this.formulaBar &&
        document.activeElement === this.input) {

        // 编辑态下保留浏览器的文本扩选行为。
        if (this.inputEditing && event.shiftKey) return;

        // 如果开启了输入框那么关闭输入框，活动的单元格调整
        if (this.inputEditing) {
            const blurEvent = new Event('blur');
            this.input.dispatchEvent(blurEvent);
        }

        let { c, r } = sheet.activeCell;
        const cell = sheet.getCell(r, c);
        const merge = cell.master ? sheet.merges.find(m => m.s.r === cell.master.r && m.s.c === cell.master.c) : null;
        const cellSide = { t: r, b: r, l: c, r: c };

        if (merge) {
            cellSide.t = merge.s.r;
            cellSide.b = merge.e.r;
            cellSide.l = merge.s.c;
            cellSide.r = merge.e.c;
        }

        if (event.key == 'Enter' || event.key == 'ArrowDown' && cellSide.b < sheet.vi.side.b) {
            sheet.activeCell.r = sheet._findNextShowRow(cellSide.b, 1);
        } else if (event.key == 'ArrowUp' && cellSide.t > sheet.vi.side.t) {
            sheet.activeCell.r = sheet._findBeforeShowRow(cellSide.t, 1);
        } else if (event.key == 'ArrowLeft' && cellSide.l > sheet.vi.side.l) {
            sheet.activeCell.c = sheet._findBeforeShowCol(cellSide.l, 1);
        } else if (event.key == 'ArrowRight' && cellSide.r < sheet.vi.side.r) {
            sheet.activeCell.c = sheet._findNextShowCol(cellSide.r, 1);
        }

        // 当活动的单元格接近边缘，调整可视区域
        const viewCol = sheet.vi.colArr;
        const viewRow = sheet.vi.rowArr;

        if (r < viewRow[2]) {
            viewRow[0] = sheet._findBeforeShowRow(viewRow[0], 1);
        } else if (r > viewRow[viewRow.length - 3]) {
            viewRow[0] = sheet._findNextShowRow(viewRow[0], 1);
        } else if (c < viewCol[2]) {
            viewCol[0] = sheet._findBeforeShowCol(viewCol[0], 1);
        } else if (c > viewCol[viewCol.length - 3]) {
            viewCol[0] = sheet._findNextShowCol(viewCol[0], 1);
        }

        sheet.activeAreas = [];
        this.r();
    }
    // Delete 删除内容
    else if (event.key == 'Delete') {
        if (sheet.hasCheckboxControl(sheet.activeAreas)) {
            event.preventDefault();
            sheet.deleteCheckboxControlOrValue(sheet.activeAreas);
            if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
                this.SN.DependencyGraph.recalculateVolatile();
            }
            this.r();
            return;
        }
        sheet.eachCells(sheet.activeAreas, (r, c) => {
            sheet.getCell(r, c).editVal = "";
        });
        // 触发 Volatile 函数重算
        if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
            this.SN.DependencyGraph.recalculateVolatile();
        }
        this.r();
    }
    // Escape 清除剪贴板状态或取消格式刷
    else if (event.key == 'Escape') {
        // 优先取消格式刷
        if (sheet.cancelBrush()) {
            this.r();
            return;
        }
        // 其次清除剪贴板
        if (sheet.hasClipboardData()) {
            sheet.clearClipboard();
            this.r();
        }
    }
    // 开启输入框条件
    else if ((event.key === ' ' || event.key === 'Spacebar') &&
        document.activeElement === this.input &&
        !this.inputEditing) {
        const active = sheet.getCell(sheet.activeCell.r, sheet.activeCell.c);
        if (isCheckboxControl(active.control)) {
            event.preventDefault();
            const areas = sheet.activeAreas?.length
                ? sheet.activeAreas
                : [{ s: sheet.activeCell, e: sheet.activeCell }];
            sheet.setCheckboxValue(areas, active.calcVal !== true);
            this.r();
            return;
        }
        this.showCellInput(true, true);
    }
    else if (document.activeElement === this.input &&
        !this.inputEditing &&
        !MODIFIER_KEYS.has(event.key)) {
        this.showCellInput(true, true);
    }
}

/**
 * Paste event handling
 * Supports HTML style parsing and image pasting
 */
export function docPaste(event) {
    // 检查焦点状态
    if (!(document.activeElement === this.input && !this.inputEditing)) {
        return;
    }

    event.preventDefault();

    // 调用 Action 层处理粘贴
    this.SN.Action.handlePasteEvent(event);
}

/**
 * Copy event handling (for blocking browser default copy)
 */
export function docCopy(event) {
    // 检查焦点状态：只有焦点在隐藏输入框且未激活编辑时才处理
    if (!(document.activeElement === this.input && !this.inputEditing)) {
        // 焦点在公式栏或其他区域，清除内部剪贴板，让系统剪贴板生效
        this.SN.Action.clearClipboard();
        return;
    }

    // 如果有选中的文本，让浏览器处理
    if (window.getSelection().toString()) {
        return;
    }

    event.preventDefault();
    this.SN.Action.copy();
}

/**
 * Cut Event Handling
 */
export function docCut(event) {
    // 检查焦点状态：只有焦点在隐藏输入框且未激活编辑时才处理
    if (!(document.activeElement === this.input && !this.inputEditing)) {
        // 焦点在公式栏或其他区域，清除内部剪贴板，让系统剪贴板生效
        this.SN.Action.clearClipboard();
        return;
    }

    // 如果有选中的文本，让浏览器处理
    if (window.getSelection().toString()) {
        return;
    }

    event.preventDefault();
    this.SN.Action.cut();
    this.r();
}
