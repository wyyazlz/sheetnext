/**
 * 键盘事件处理模块
 * 调用 Action 层的剪贴板方法
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

export function docKeyDown(event) {
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

    // F9 - 全局重算（包括 Volatile 函数）
    if (event.key === 'F9') {
        event.preventDefault();
        this.SN.recalculate(event.shiftKey); // Shift+F9 仅重算当前表
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
        } else if ((event.key == 'ArrowRight' || event.key == 'Tab') && cellSide.r < sheet.vi.side.r) {
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
    else if (document.activeElement === this.input && !this.inputEditing) {
        this.showCellInput(true, true);
    }
}

/**
 * 粘贴事件处理
 * 支持 HTML 样式解析和图片粘贴
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
 * 复制事件处理（用于拦截浏览器默认复制）
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
 * 剪切事件处理
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
