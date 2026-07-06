/**
 * Inline Formula Editor Controller
 * Replicates Excel's formula entry experience on top of the cell editor:
 * - Colors cell/range references inside the contenteditable editor
 * - Draws matching colored frames around referenced areas on the canvas
 * - Function name autocomplete dropdown fed by FormulaCatalog
 * - Parameter hint showing the current argument of the innermost call
 * - Reference pointing: click/drag on the grid or arrow keys insert refs
 * - F4 cycles absolute/relative anchors of the reference at the cursor
 */
import { analyzeFormulaInput, cycleReferenceAnchor } from '../Formula/FormulaEditContext.js';
import { FormulaCatalog } from '../Formula/FormulaCatalog.js';
import { getFormulaSignature } from '../../assets/formulaSignature.js';

// Reference colors, cycled per distinct reference (aligned with Excel's palette)
const REF_COLORS = ['#2E75B6', '#C00000', '#7030A0', '#548235', '#BF8F00', '#0E8A8A', '#C55A9B'];
const AC_MAX_ITEMS = 100;
const ARROW_DELTAS = {
    ArrowUp: { r: -1, c: 0 },
    ArrowDown: { r: 1, c: 0 },
    ArrowLeft: { r: 0, c: -1 },
    ArrowRight: { r: 0, c: 1 }
};

// Serialize a DOM subtree to plain text, <br> counts as '\n'
function nodeText(node) {
    if (node.nodeType === 3) return node.textContent;
    if (node.nodeType === 1 && node.tagName === 'BR') return '\n';
    let out = '';
    for (let child = node.firstChild; child; child = child.nextSibling) {
        out += nodeText(child);
    }
    return out;
}

// Append text to a parent element, converting '\n' to <br>
function appendTextWithBreaks(parent, text) {
    const lines = String(text).split('\n');
    lines.forEach((line, i) => {
        if (i > 0) parent.appendChild(document.createElement('br'));
        if (line) parent.appendChild(document.createTextNode(line));
    });
}

export default class FormulaEditor {
    /**
     * @param {Canvas} canvas - Canvas instance hosting the cell editor
     */
    constructor(canvas) {
        /**
         * Canvas instance
         * @type {Canvas}
         */
        this.canvas = canvas;
        /**
         * SheetNext Main Instance
         * @type {Object}
         */
        this.SN = canvas.SN;
        this._active = false;
        this._analysis = null;
        this._refAreas = [];
        this._refSignature = '';
        this._acItems = [];
        this._acIndex = 0;
        this._acPrefix = null;
        this._acSuppressed = false;
        this._pending = null;
        this._pointAnchor = null;
        this._pointCursor = null;
        this._listEl = null;
        this._hintEl = null;
        this._composing = false;
        this._selectionRaf = null;

        canvas.input.addEventListener('compositionstart', () => { this._composing = true; });
        canvas.input.addEventListener('compositionend', () => {
            this._composing = false;
            this.onInput();
        });
        document.addEventListener('selectionchange', () => this._onSelectionChange());
    }

    /**
     * Whether the formula editing mode is currently active
     * @type {boolean}
     */
    get active() {
        return this._active === true;
    }

    /**
     * Start a fresh editing session, called when the cell editor opens.
     * @returns {void}
     */
    beginSession() {
        this._pending = null;
        this._pointAnchor = null;
        this._pointCursor = null;
        this._acSuppressed = false;
        this._acIndex = 0;
        this.syncFromInput();
    }

    /**
     * Re-read the editor content and refresh coloring, overlays and highlights.
     * @returns {void}
     */
    syncFromInput() {
        if (!this.canvas.inputEditing) {
            if (this._active) this.deactivate();
            return;
        }
        const text = this._getText();
        if (!text.startsWith('=')) {
            this._deactivateInline(text);
            return;
        }
        this._active = true;
        this.canvas.input.classList.add('sn-formula-mode');
        this._render(text, this._getCaret());
    }

    /**
     * User typed into the cell editor.
     * @returns {void}
     */
    onInput() {
        if (!this.canvas.inputEditing) return;
        if (this._composing) return;
        this._pending = null;
        this._pointAnchor = null;
        this._pointCursor = null;
        this._acSuppressed = false;
        this.syncFromInput();
    }

    /**
     * Editor content was mirrored from the formula bar, refresh without caret handling.
     * @returns {void}
     */
    onExternalTextChange() {
        if (!this.canvas.inputEditing) return;
        this._pending = null;
        this._pointAnchor = null;
        this._pointCursor = null;
        const text = this._getText();
        if (!text.startsWith('=')) {
            this._deactivateInline(text);
            return;
        }
        this._active = true;
        this.canvas.input.classList.add('sn-formula-mode');
        const analysis = analyzeFormulaInput(text, null);
        this._renderTokens(text, analysis.refs);
        this._analysis = analysis;
        this._hideDropdown();
        this._hideHint();
        this._updateRefAreas(analysis);
    }

    /**
     * Intercept keyboard events during formula editing.
     * Handles dropdown navigation, Tab-accept, Escape, F4 and arrow-key pointing.
     * @param {KeyboardEvent} event - Keydown event from the document handler
     * @returns {boolean} true when the event was consumed
     */
    handleKeyDown(event) {
        const canvas = this.canvas;
        if (!this._active || !canvas.inputEditing) return false;
        if (document.activeElement !== canvas.input) return false;
        if (event.ctrlKey || event.metaKey || event.altKey) return false;

        const key = event.key;

        if (this._isDropdownVisible()) {
            if (key === 'ArrowDown' || key === 'ArrowUp') {
                event.preventDefault();
                this._moveAcIndex(key === 'ArrowDown' ? 1 : -1);
                return true;
            }
            if (key === 'Tab') {
                event.preventDefault();
                this._acceptAutocomplete(this._acIndex);
                return true;
            }
            if (key === 'Escape') {
                event.preventDefault();
                this._acSuppressed = true;
                this._hideDropdown();
                return true;
            }
        }

        if (key === 'F4') {
            if (this._applyF4()) {
                event.preventDefault();
                return true;
            }
            return false;
        }

        if (key === 'Escape') {
            event.preventDefault();
            this._cancelEdit();
            return true;
        }

        if (ARROW_DELTAS[key]) {
            if (this._pointByKey(key, event.shiftKey)) {
                event.preventDefault();
                return true;
            }
            return false;
        }

        return false;
    }

    /**
     * Intercept a canvas mousedown during formula editing.
     * When the cursor is at a reference-insertable position the click/drag
     * inserts a cell/range/row/column reference instead of committing the edit.
     * @param {MouseEvent} event - Mousedown event on the handle layer
     * @returns {boolean} true when the event was consumed
     */
    handleCanvasMousedown(event) {
        const canvas = this.canvas;
        if (!this._active || !canvas.inputEditing) return false;
        if (event.button !== 0) return false;
        if (document.activeElement !== canvas.input) return false;

        const caret = this._getCaret();
        if (caret == null) return false;
        const text = this._getText();
        let range = null;
        if (this._pending && caret === this._pending.end) {
            range = { ...this._pending };
        } else {
            const analysis = analyzeFormulaInput(text, caret);
            if (!analysis) return false;
            if (analysis.refAtCursor) {
                range = { start: analysis.refAtCursor.start, end: analysis.refAtCursor.end };
            } else if (analysis.canInsertRef) {
                range = { start: caret, end: caret };
            } else {
                return false; // 不可插入位置：走默认提交逻辑（与 Excel 一致）
            }
        }

        event.preventDefault(); // 阻止焦点转移，编辑器不失焦不提交

        const sheet = canvas.activeSheet;
        const { x, y } = canvas.getEventPosition(event);
        const inColHead = y < sheet.headHeight;
        const inRowHead = x < sheet.indexWidth;
        if (inColHead && inRowHead) return true; // 全选角落：忽略

        const mode = inColHead ? 'col' : inRowHead ? 'row' : 'cell';
        const idx = canvas.getCellIndex(event);
        const anchor = { r: Math.max(0, idx.r), c: Math.max(0, idx.c) };

        const apply = (target) => {
            const refText = this._refTextBetween(mode, anchor, target);
            this._replaceRange(range.start, range.end, refText);
            range.end = range.start + refText.length;
            this._pending = { start: range.start, end: range.end };
        };
        apply(anchor);
        this._pointAnchor = anchor;
        this._pointCursor = { ...anchor };

        // 拖拽扩展为区域引用
        canvas.disableMoveCheck = true;
        const layer = canvas.handleLayer;
        const onMove = (moveEvent) => {
            const cur = canvas.getCellIndex(moveEvent);
            const target = { r: Math.max(0, cur.r), c: Math.max(0, cur.c) };
            if (target.r === this._pointCursor.r && target.c === this._pointCursor.c) return;
            this._pointCursor = target;
            apply(target);
        };
        layer.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', () => {
            layer.removeEventListener('mousemove', onMove);
            canvas.disableMoveCheck = false;
            canvas.input.focus();
        }, { once: true });
        return true;
    }

    /**
     * Draw colored frames around referenced areas, called from the render pipeline.
     * @param {Sheet} sheet - Sheet being rendered
     * @returns {void}
     */
    drawRefHighlights(sheet) {
        if (!this._active || !this._refAreas.length) return;
        const canvas = this.canvas;
        const bc = canvas.bc;
        if (!bc) return;
        const clipX = sheet.indexWidth;
        const clipY = sheet.headHeight;
        const clipW = canvas.viewWidth - clipX;
        const clipH = canvas.viewHeight - clipY;
        if (clipW <= 0 || clipH <= 0) return;

        const sheetName = String(sheet.name ?? '').toLowerCase();
        bc.save();
        bc.beginPath();
        bc.rect(clipX, clipY, clipW, clipH);
        bc.clip();
        for (const ref of this._refAreas) {
            if (String(ref.sheetName ?? '').toLowerCase() !== sheetName) continue;
            const info = sheet.getAreaInviewInfo({ s: ref.s, e: ref.e });
            if (!info.inView) continue;
            const color = REF_COLORS[ref.colorIndex % REF_COLORS.length];
            const rect = canvas.snapRect(info.x, info.y, info.w, info.h);
            bc.fillStyle = color + '26';
            bc.fillRect(rect.x, rect.y, rect.w, rect.h);
            bc.strokeStyle = color;
            bc.lineWidth = 1.5;
            bc.strokeRect(rect.x, rect.y, rect.w, rect.h);
            // 四角小方块（Excel 风格拖拽柄外观）
            bc.fillStyle = color;
            const size = 4;
            [[rect.x, rect.y], [rect.x + rect.w, rect.y], [rect.x, rect.y + rect.h], [rect.x + rect.w, rect.y + rect.h]]
                .forEach(([cx, cy]) => bc.fillRect(cx - size / 2, cy - size / 2, size, size));
        }
        bc.lineWidth = canvas.px(1);
        bc.restore();
    }

    /**
     * Reposition editor overlays after view scrolling during formula editing.
     * @returns {void}
     */
    onViewScrolled() {
        if (!this._active) return;
        this.canvas.updateInputPosition();
        if (this._isDropdownVisible()) this._positionOverlay(this._listEl);
        if (this._hintEl && this._hintEl.style.display !== 'none') this._positionOverlay(this._hintEl);
    }

    /**
     * Leave formula mode and clear overlays/highlights.
     * @param {boolean} [render=true] - Whether to trigger a snapshot repaint
     * @returns {void}
     */
    deactivate(render = true) {
        this._active = false;
        this._pending = null;
        this._pointAnchor = null;
        this._pointCursor = null;
        this._analysis = null;
        this.canvas.input.classList.remove('sn-formula-mode');
        this._hideDropdown();
        this._hideHint();
        if (this._refAreas.length) {
            this._refAreas = [];
            this._refSignature = '';
            if (render) this.canvas.r('s');
        }
    }

    // ==================== 内部：状态与渲染 ====================

    // Leave formula mode while editing continues (text no longer starts with '=')
    _deactivateInline(text) {
        if (!this._active) return;
        const input = this.canvas.input;
        this._active = false;
        this._pending = null;
        this._pointAnchor = null;
        this._pointCursor = null;
        this._analysis = null;
        input.classList.remove('sn-formula-mode');
        this._hideDropdown();
        this._hideHint();
        // 拉平残留的引用着色 span，避免被富文本提交路径误判
        if (input.querySelector('.sn-formula-ref')) {
            const caret = this._getCaret();
            this._renderTokens(text, []);
            if (caret != null && document.activeElement === input) this._setCaret(caret);
        }
        if (this._refAreas.length) {
            this._refAreas = [];
            this._refSignature = '';
            this.canvas.r('s');
        }
    }

    // Full refresh: rebuild colored DOM, restore caret, sync overlays and highlights
    _render(text, caret) {
        const analysis = analyzeFormulaInput(text, caret);
        if (!analysis) return;
        if (!this._composing) {
            this._renderTokens(text, analysis.refs);
            // 仅在编辑器持有焦点时恢复光标，避免从公式栏抢占焦点
            if (caret != null && document.activeElement === this.canvas.input) {
                this._setCaret(caret);
            }
        }
        this.canvas.formulaBar.value = text;
        this._analysis = analysis;
        this._updateDropdown(analysis);
        this._updateHint(analysis);
        this._updateRefAreas(analysis);
    }

    // Rebuild editor children: plain text nodes + colored spans for references
    _renderTokens(text, refs) {
        const input = this.canvas.input;
        const frag = document.createDocumentFragment();
        let pos = 0;
        const pushPlain = (to) => {
            if (to <= pos) return;
            appendTextWithBreaks(frag, text.slice(pos, to));
            pos = to;
        };
        for (const ref of refs) {
            pushPlain(ref.start);
            const span = document.createElement('span');
            span.className = 'sn-formula-ref';
            span.style.color = REF_COLORS[ref.colorIndex % REF_COLORS.length];
            appendTextWithBreaks(span, text.slice(ref.start, ref.end));
            frag.appendChild(span);
            pos = ref.end;
        }
        pushPlain(text.length);
        // 末尾换行需要占位 <br> 才能在 contenteditable 中可见
        if (text.endsWith('\n')) frag.appendChild(document.createElement('br'));
        input.replaceChildren(frag);
    }

    // Replace a text range and re-render with the caret after the inserted text
    _replaceRange(start, end, insert) {
        const text = this._getText();
        const next = text.slice(0, start) + insert + text.slice(end);
        this._render(next, start + insert.length);
    }

    _getText() {
        const text = nodeText(this.canvas.input);
        return text.endsWith('\n') ? text.slice(0, -1) : text;
    }

    _getCaret() {
        const selection = window.getSelection();
        if (!selection || !selection.rangeCount) return null;
        const range = selection.getRangeAt(0);
        if (!this.canvas.input.contains(range.endContainer)) return null;
        const pre = document.createRange();
        pre.selectNodeContents(this.canvas.input);
        pre.setEnd(range.endContainer, range.endOffset);
        return nodeText(pre.cloneContents()).length;
    }

    _setCaret(offset) {
        const selection = window.getSelection();
        if (!selection) return;
        const range = document.createRange();
        let remaining = offset;
        let done = false;
        const walk = (node) => {
            if (done) return;
            if (node.nodeType === 3) {
                const len = node.textContent.length;
                if (remaining <= len) {
                    range.setStart(node, remaining);
                    done = true;
                } else {
                    remaining -= len;
                }
                return;
            }
            if (node.nodeType === 1 && node.tagName === 'BR') {
                if (remaining <= 0) {
                    range.setStartBefore(node);
                    done = true;
                } else {
                    remaining -= 1;
                }
                return;
            }
            for (let child = node.firstChild; child && !done; child = child.nextSibling) walk(child);
        };
        walk(this.canvas.input);
        if (!done) {
            range.selectNodeContents(this.canvas.input);
            range.collapse(false);
        } else {
            range.collapse(true);
        }
        selection.removeAllRanges();
        selection.addRange(range);
    }

    _onSelectionChange() {
        if (!this._active || this._composing) return;
        if (document.activeElement !== this.canvas.input) return;
        if (this._selectionRaf) return;
        this._selectionRaf = requestAnimationFrame(() => {
            this._selectionRaf = null;
            if (!this._active || this._composing) return;
            if (document.activeElement !== this.canvas.input) return;
            const analysis = analyzeFormulaInput(this._getText(), this._getCaret());
            if (!analysis) return;
            this._analysis = analysis;
            this._updateDropdown(analysis);
            this._updateHint(analysis);
        });
    }

    // ==================== 内部：引用高亮 ====================

    _updateRefAreas(analysis) {
        const formula = this.SN.Formula;
        const sheet = this.canvas.activeSheet;
        const areas = [];
        const seen = new Set();
        if (formula && sheet) {
            for (const ref of analysis.refs) {
                let parsed;
                try {
                    parsed = formula.parseDependencyAreas(ref.text, sheet.name);
                } catch {
                    continue; // 输入到一半的引用，忽略
                }
                const area = parsed?.[0];
                if (!area) continue;
                const key = `${area.sheetName}|${area.s.r},${area.s.c},${area.e.r},${area.e.c}|${ref.colorIndex}`;
                if (seen.has(key)) continue;
                seen.add(key);
                areas.push({ sheetName: area.sheetName, s: area.s, e: area.e, colorIndex: ref.colorIndex });
            }
        }
        const signature = [...seen].join(';');
        if (signature !== this._refSignature) {
            this._refSignature = signature;
            this._refAreas = areas;
            this.canvas.r('s');
        }
    }

    // ==================== 内部：函数联想下拉 ====================

    _ensureDom() {
        const input = this.canvas.input;
        const host = input.offsetParent || input.parentElement;
        if (!host) return;
        if (!this._listEl || !this._listEl.isConnected) {
            const list = document.createElement('div');
            list.className = 'sn-formula-ac sn-scrollbar';
            list.style.display = 'none';
            list.addEventListener('mousedown', (event) => {
                event.preventDefault(); // 保持编辑器焦点
                const item = event.target.closest('.sn-formula-ac-item');
                if (item) this._acceptAutocomplete(Number(item.dataset.index));
            });
            host.appendChild(list);
            this._listEl = list;
        }
        if (!this._hintEl || !this._hintEl.isConnected) {
            const hint = document.createElement('div');
            hint.className = 'sn-formula-hint';
            hint.style.display = 'none';
            hint.addEventListener('mousedown', (event) => event.preventDefault());
            host.appendChild(hint);
            this._hintEl = hint;
        }
    }

    _updateDropdown(analysis) {
        const prefix = analysis?.fnPrefix;
        if (!prefix || this._acSuppressed || document.activeElement !== this.canvas.input) {
            this._hideDropdown();
            return;
        }
        const term = prefix.text.toUpperCase();
        const items = FormulaCatalog.getFormulasByCategory('all')
            .filter(item => item.name.startsWith(term))
            .slice(0, AC_MAX_ITEMS);
        if (!items.length) {
            this._hideDropdown();
            return;
        }
        this._ensureDom();
        if (!this._listEl) return;
        this._acItems = items;
        this._acPrefix = { ...prefix };
        this._acIndex = 0;
        const frag = document.createDocumentFragment();
        items.forEach((item, i) => {
            const el = document.createElement('div');
            el.className = 'sn-formula-ac-item' + (i === 0 ? ' active' : '');
            el.dataset.index = String(i);
            const name = document.createElement('span');
            name.className = 'sn-formula-ac-name';
            name.textContent = item.name;
            const desc = document.createElement('span');
            desc.className = 'sn-formula-ac-desc';
            desc.textContent = item.desc || '';
            el.append(name, desc);
            frag.appendChild(el);
        });
        this._listEl.replaceChildren(frag);
        this._listEl.style.display = 'block';
        this._listEl.scrollTop = 0;
        this._positionOverlay(this._listEl);
        this._hideHint(); // 下拉与参数提示互斥，下拉优先
    }

    _hideDropdown() {
        if (this._listEl) this._listEl.style.display = 'none';
        this._acItems = [];
        this._acPrefix = null;
    }

    _isDropdownVisible() {
        return !!this._listEl && this._listEl.style.display !== 'none' && this._acItems.length > 0;
    }

    _moveAcIndex(step) {
        if (!this._acItems.length) return;
        const next = (this._acIndex + step + this._acItems.length) % this._acItems.length;
        this._acIndex = next;
        this._listEl.querySelectorAll('.sn-formula-ac-item').forEach((el, i) => {
            el.classList.toggle('active', i === next);
            if (i === next) el.scrollIntoView({ block: 'nearest' });
        });
    }

    _acceptAutocomplete(index) {
        const item = this._acItems[index];
        const prefix = this._acPrefix;
        if (!item || !prefix) return;
        FormulaCatalog.addRecent(item.name);
        this._hideDropdown();
        this._replaceRange(prefix.start, prefix.end, item.name + '(');
    }

    // ==================== 内部：参数签名提示 ====================

    _updateHint(analysis) {
        if (this._isDropdownVisible()) {
            this._hideHint();
            return;
        }
        const call = analysis?.call;
        if (!call || document.activeElement !== this.canvas.input) {
            this._hideHint();
            return;
        }
        const args = getFormulaSignature(call.name);
        if (!args) {
            this._hideHint();
            return;
        }
        this._ensureDom();
        if (!this._hintEl) return;
        const current = args.length ? Math.min(call.argIndex, args.length - 1) : -1;
        const frag = document.createDocumentFragment();
        const name = document.createElement('span');
        name.className = 'sn-formula-hint-name';
        name.textContent = call.name;
        frag.appendChild(name);
        frag.appendChild(document.createTextNode('('));
        args.forEach((arg, i) => {
            if (i > 0) frag.appendChild(document.createTextNode(', '));
            const el = document.createElement(i === current ? 'b' : 'span');
            el.textContent = arg;
            frag.appendChild(el);
        });
        frag.appendChild(document.createTextNode(')'));
        this._hintEl.replaceChildren(frag);
        this._hintEl.style.display = 'block';
        this._positionOverlay(this._hintEl);
    }

    _hideHint() {
        if (this._hintEl) this._hintEl.style.display = 'none';
    }

    // Place an overlay right below the cell editor, flipping above when overflowing
    _positionOverlay(el) {
        if (!el) return;
        const input = this.canvas.input;
        el.style.left = `${input.offsetLeft}px`;
        el.style.top = `${input.offsetTop + input.offsetHeight + 2}px`;
        const host = el.offsetParent;
        if (!host) return;
        if (el.offsetTop + el.offsetHeight > host.clientHeight) {
            el.style.top = `${Math.max(0, input.offsetTop - el.offsetHeight - 2)}px`;
        }
        if (el.offsetLeft + el.offsetWidth > host.clientWidth) {
            el.style.left = `${Math.max(0, host.clientWidth - el.offsetWidth - 2)}px`;
        }
    }

    // ==================== 内部：引用插入（Point 模式）与 F4 ====================

    // Arrow-key pointing: insert/move a reference like Excel's Enter mode
    _pointByKey(key, shift) {
        const canvas = this.canvas;
        const sheet = canvas.activeSheet;
        if (!sheet) return false;
        const caret = this._getCaret();
        if (caret == null) return false;
        const text = this._getText();
        const delta = ARROW_DELTAS[key];

        let range = null;
        let base = null;
        if (this._pending && caret === this._pending.end && this._pointCursor) {
            range = { ...this._pending };
            base = this._pointCursor;
        } else {
            this._pending = null;
            this._pointAnchor = null;
            this._pointCursor = null;
            const analysis = analyzeFormulaInput(text, caret);
            if (!analysis) return false;
            if (analysis.refAtCursor) {
                base = this._refStartCell(analysis.refAtCursor.text);
                if (!base) return false;
                range = { start: analysis.refAtCursor.start, end: analysis.refAtCursor.end };
            } else if (analysis.canInsertRef) {
                base = { r: canvas.currentCell.r, c: canvas.currentCell.c };
                range = { start: caret, end: caret };
            } else {
                return false;
            }
        }

        const next = {
            r: Math.min(Math.max(base.r + delta.r, 0), sheet.rowCount - 1),
            c: Math.min(Math.max(base.c + delta.c, 0), sheet.colCount - 1)
        };
        const anchor = shift ? (this._pointAnchor ?? next) : next;
        const refText = this._refTextBetween('cell', anchor, next);
        this._replaceRange(range.start, range.end, refText);
        this._pending = { start: range.start, end: range.start + refText.length };
        this._pointAnchor = anchor;
        this._pointCursor = next;
        return true;
    }

    // Resolve the start cell of a same-sheet reference, null when not pointable
    _refStartCell(refText) {
        const formula = this.SN.Formula;
        const sheet = this.canvas.activeSheet;
        if (!formula || !sheet) return null;
        try {
            const area = formula.parseDependencyAreas(refText, sheet.name)?.[0];
            if (!area) return null;
            if (String(area.sheetName ?? '').toLowerCase() !== String(sheet.name ?? '').toLowerCase()) return null;
            return { r: area.s.r, c: area.s.c };
        } catch {
            return null;
        }
    }

    // Build the reference text between two cells for the given pointing mode
    _refTextBetween(mode, a, b) {
        const Utils = this.SN.Utils;
        if (mode === 'col') {
            const c1 = Math.min(a.c, b.c);
            const c2 = Math.max(a.c, b.c);
            return `${Utils.numToChar(c1)}:${Utils.numToChar(c2)}`;
        }
        if (mode === 'row') {
            const r1 = Math.min(a.r, b.r);
            const r2 = Math.max(a.r, b.r);
            return `${r1 + 1}:${r2 + 1}`;
        }
        if (a.r === b.r && a.c === b.c) return Utils.cellNumToStr(a);
        return Utils.rangeNumToStr({
            s: { r: Math.min(a.r, b.r), c: Math.min(a.c, b.c) },
            e: { r: Math.max(a.r, b.r), c: Math.max(a.c, b.c) }
        });
    }

    // F4: cycle absolute/relative anchors of the reference at the cursor
    _applyF4() {
        const caret = this._getCaret();
        if (caret == null) return false;
        const text = this._getText();
        let target = null;
        if (this._pending && caret === this._pending.end) {
            target = { ...this._pending };
        } else {
            const analysis = analyzeFormulaInput(text, caret);
            if (analysis?.refAtCursor) {
                target = { start: analysis.refAtCursor.start, end: analysis.refAtCursor.end };
            }
        }
        if (!target) return false;
        const current = text.slice(target.start, target.end);
        const cycled = cycleReferenceAnchor(current);
        if (cycled === current) return false;
        this._replaceRange(target.start, target.end, cycled);
        if (this._pending && this._pending.start === target.start) {
            this._pending = { start: target.start, end: target.start + cycled.length };
        }
        return true;
    }

    // Escape: cancel the edit without committing (formula mode only)
    _cancelEdit() {
        const canvas = this.canvas;
        this.deactivate(false);
        canvas.inputEditing = false;
        canvas.input.innerHTML = '';
        canvas.input.focus();
        canvas.r('s'); // 恢复选区显示与公式栏内容
    }
}
