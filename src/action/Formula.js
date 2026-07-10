import { FormulaCatalog } from '../core/Formula/FormulaCatalog.js';
import { getFormulaSignature } from '../assets/formulaSignature.js';
import getSvg from '../assets/mainSvgs.js';

function escapeHTML(value) {
    return String(value ?? '').replace(/[&<>'"]/g, char => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        "'": '&#39;',
        '"': '&quot;'
    })[char]);
}

function cloneAreas(areas) {
    return (areas || []).map(area => ({
        s: { r: area.s.r, c: area.s.c },
        e: { r: area.e.r, c: area.e.c }
    }));
}

function getNodeTextLength(node) {
    if (!node) return 0;
    if (node.nodeType === 3) return node.textContent?.length || 0;
    if (node.nodeType === 1 && node.tagName === 'BR') return 1;
    let length = 0;
    for (let child = node.firstChild; child; child = child.nextSibling) {
        length += getNodeTextLength(child);
    }
    return length;
}

function getFormulaSelection(canvas, draft) {
    const fallback = { start: draft.length, end: draft.length };
    if (!canvas) return fallback;
    if (document.activeElement === canvas.formulaBar) {
        return {
            start: canvas.formulaBar.selectionStart ?? draft.length,
            end: canvas.formulaBar.selectionEnd ?? draft.length
        };
    }
    if (document.activeElement !== canvas.input) return fallback;
    const selection = window.getSelection();
    if (!selection?.rangeCount) return fallback;
    const range = selection.getRangeAt(0);
    if (!canvas.input.contains(range.startContainer) || !canvas.input.contains(range.endContainer)) return fallback;
    const offsetAt = (container, offset) => {
        const before = document.createRange();
        before.selectNodeContents(canvas.input);
        before.setEnd(container, offset);
        return getNodeTextLength(before.cloneContents());
    };
    return {
        start: offsetAt(range.startContainer, range.startOffset),
        end: offsetAt(range.endContainer, range.endOffset)
    };
}

function captureFormulaInsertContext(SN) {
    const canvas = SN.Canvas;
    const targetSheet = SN.activeSheet;
    const editingCell = canvas?.inputEditing && canvas.currentCell
        ? canvas.currentCell
        : targetSheet.activeCell;
    const formulaBarActive = document.activeElement === canvas?.formulaBar;
    const draft = canvas?.inputEditing
        ? (formulaBarActive ? canvas.formulaBar.value : canvas.input.innerText) || ''
        : '';
    const selection = getFormulaSelection(canvas, draft);
    return {
        targetSheet,
        targetCell: { r: editingCell.r, c: editingCell.c },
        activeAreas: cloneAreas(targetSheet.activeAreas),
        wasEditing: Boolean(canvas?.inputEditing),
        draft,
        selectionStart: selection.start,
        selectionEnd: selection.end
    };
}

function suspendFormulaEdit(SN) {
    const canvas = SN.Canvas;
    if (!canvas) return;
    canvas._suspendEditCommit = true;
    if (!canvas.inputEditing) return;
    canvas.formulaEditor?.deactivate(false);
    canvas.inputEditing = false;
    canvas.input.innerHTML = '';
}

function releaseFormulaEditSuspension(SN) {
    if (SN.Canvas) SN.Canvas._suspendEditCommit = false;
}

function focusOpenArgumentDialog(SN) {
    const modal = SN.containerDom.querySelector('.sn-formula-args-modal');
    const target = modal?.querySelector('.sn-formula-arg-row.active .sn-formula-arg-input')
        || modal?.querySelector('.sn-formula-arg-input')
        || modal?.querySelector('.sn-modal-btn-confirm');
    target?.focus();
}

function restoreFormulaTarget(SN, context, restoreDraft = false, fullRender = false) {
    const canvas = SN.Canvas;
    const sheet = context.targetSheet;
    sheet.activeCell = { ...context.targetCell };
    sheet.activeAreas = cloneAreas(context.activeAreas);
    if (SN.activeSheet !== sheet) {
        SN.activeSheet = sheet;
    } else {
        fullRender ? canvas?.r() : canvas?.r('s');
    }

    if (!restoreDraft || !context.wasEditing || !canvas) {
        canvas?.input?.focus();
        return;
    }
    canvas.showCellInput(false, false);
    canvas.input.innerText = context.draft;
    canvas.formulaBar.value = context.draft;
    canvas.formulaEditor?.beginSession();
    canvas.input.focus();
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.selectNodeContents(canvas.input);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
}

function getArgumentSpec(name, SN) {
    const signature = getFormulaSignature(name);
    if (signature === null) {
        return {
            allowMore: true,
            definitions: [{
                label: SN.t('action.formula.content.arguments.argument', { index: 1 }),
                optional: false
            }]
        };
    }
    return {
        allowMore: signature.includes('...'),
        definitions: signature
            .filter(item => item !== '...')
            .map((item) => {
                const optional = item.startsWith('[') && item.endsWith(']');
                return {
                    label: optional ? item.slice(1, -1) : item,
                    optional
                };
            })
    };
}

function buildArgumentRowHTML(definition, index, SN) {
    const label = escapeHTML(definition.label);
    const selectRangeText = escapeHTML(SN.t('action.formula.content.arguments.selectRange'));
    const placeholder = escapeHTML(SN.t('action.formula.content.arguments.inputPlaceholder'));
    const optional = definition.optional
        ? `<span class="sn-formula-arg-optional">${escapeHTML(SN.t('action.formula.content.arguments.optional'))}</span>`
        : '';
    return `
        <div class="sn-formula-arg-row" data-arg-index="${index}">
            <div class="sn-formula-arg-label" title="${label}">
                <span>${label}</span>${optional}
            </div>
            <div class="sn-formula-arg-control">
                <input type="text" class="sn-formula-arg-input" data-arg-index="${index}"
                    placeholder="${placeholder}" aria-label="${label}">
                <button type="button" class="sn-formula-range-btn" data-arg-index="${index}"
                    title="${selectRangeText}" aria-label="${selectRangeText}">
                    ${getSvg('sort_moveUp')}
                </button>
            </div>
        </div>
    `;
}

function buildArgumentDialogHTML(name, spec, SN) {
    const rows = spec.definitions.length
        ? spec.definitions.map((definition, index) => buildArgumentRowHTML(definition, index, SN)).join('')
        : `<div class="sn-formula-args-empty">${escapeHTML(SN.t('action.formula.content.arguments.noArguments'))}</div>`;
    const addButton = spec.allowMore
        ? `<button type="button" class="sn-formula-add-arg">
                ${getSvg('plus')}
                <span>${escapeHTML(SN.t('action.formula.content.arguments.addArgument'))}</span>
           </button>`
        : '';
    return `
        <div class="sn-formula-args">
            <div class="sn-formula-args-name">${escapeHTML(name)}</div>
            <div class="sn-formula-args-list sn-scrollbar">${rows}</div>
            ${addButton}
        </div>
    `;
}

function buildDialogHTML(categories, activeKey, SN) {
    const options = categories.map(cat =>
        `<option value="${cat.key}"${cat.key === activeKey ? ' selected' : ''}>${SN.t(cat.labelKey)}</option>`
    ).join('');

    return `
        <div class="sn-formula-dialog">
            <div class="sn-formula-row">
                <label>${SN.t('action.formula.content.searchLabel')}</label>
                <input type="text" class="sn-formula-search-input" placeholder="${SN.t('action.formula.content.searchPlaceholder')}">
            </div>
            <div class="sn-formula-row">
                <label>${SN.t('action.formula.content.categoryLabel')}</label>
                <select class="sn-formula-category-select">${options}</select>
            </div>
            <div class="sn-formula-row sn-formula-row-list">
                <label>${SN.t('action.formula.content.functionLabel')}</label>
                <div class="sn-formula-list"></div>
            </div>
            <div class="sn-formula-detail">
                <div class="sn-formula-detail-title"></div>
                <div class="sn-formula-detail-desc"></div>
                <div class="sn-formula-detail-meta"></div>
            </div>
        </div>
    `;
}

function renderFormulaList(listEl, formulas, activeName, SN) {
    if (!formulas.length) {
        listEl.innerHTML = `<div class="sn-formula-empty">${SN.t('action.formula.content.empty')}</div>`;
        return;
    }
    listEl.innerHTML = formulas.map(item =>
        `<div class="sn-formula-item${item.name === activeName ? ' active' : ''}" data-name="${item.name}">${item.name}</div>`
    ).join('');
}

function renderFormulaDetail(bodyEl, formula, SN) {
    const titleEl = bodyEl.querySelector('.sn-formula-detail-title');
    const descEl = bodyEl.querySelector('.sn-formula-detail-desc');
    const metaEl = bodyEl.querySelector('.sn-formula-detail-meta');

    if (!formula) {
        titleEl.textContent = '';
        descEl.textContent = '';
        metaEl.textContent = '';
        return;
    }

    titleEl.textContent = `${formula.name}(...)`;
    descEl.textContent = /[\u4e00-\u9fff]/.test(formula.desc || '') ? SN.t('action.formula.content.noDescription') : (formula.desc || '');
    metaEl.textContent = [
        formula.categoryLabelKey && `${SN.t('action.formula.content.detail.category')}: ${SN.t(formula.categoryLabelKey)}`,
        formula.version && `${SN.t('action.formula.content.detail.version')}: ${formula.version}`
    ].filter(Boolean).join(' · ');
}

function openFunctionArguments(name, context) {
    const SN = this.SN;
    const canvas = SN.Canvas;
    const spec = getArgumentSpec(name, SN);
    let stopRangePicker = () => {};
    let activeInput = null;
    let pendingReference = '';
    let committed = false;
    this._formulaArgumentDialogOpen = true;

    if (spec.definitions.length > 0) {
        stopRangePicker = canvas.startRangePicker({
            targetSheetName: context.targetSheet.name,
            onChange: (reference) => {
                pendingReference = reference;
                if (!activeInput?.isConnected) return;
                activeInput.value = reference;
                activeInput.dispatchEvent(new Event('input', { bubbles: true }));
            }
        });
    }

    SN.Utils.modal({
        titleKey: 'action.formula.modal.m002.title',
        title: 'Function Arguments',
        maxWidth: '460px',
        modalClass: 'sn-formula-args-modal',
        bodyClass: 'sn-modal-body-formula-args',
        modeless: spec.definitions.length > 0,
        confirmKey: 'action.formula.modal.m002.confirm',
        confirmText: 'OK',
        cancelKey: 'action.formula.modal.m002.cancel',
        cancelText: 'Cancel',
        content: buildArgumentDialogHTML(name, spec, SN),
        onOpen: (bodyEl) => {
            suspendFormulaEdit(SN);
            const argsEl = bodyEl.querySelector('.sn-formula-args');
            const listEl = bodyEl.querySelector('.sn-formula-args-list');
            const addButton = bodyEl.querySelector('.sn-formula-add-arg');
            const modalEl = bodyEl.closest('.sn-modal');
            activeInput = listEl?.querySelector('.sn-formula-arg-input') || null;

            const setActiveInput = (input) => {
                if (!input) return;
                activeInput = input;
                listEl.querySelectorAll('.sn-formula-arg-row').forEach(row => {
                    row.classList.toggle('active', row === input.closest('.sn-formula-arg-row'));
                });
            };

            listEl?.addEventListener('focusin', (event) => {
                const input = event.target.closest('.sn-formula-arg-input');
                if (input) setActiveInput(input);
            });
            listEl?.addEventListener('mousedown', (event) => {
                if (event.target.closest('.sn-formula-range-btn')) event.preventDefault();
            });
            listEl?.addEventListener('click', (event) => {
                const button = event.target.closest('.sn-formula-range-btn');
                if (!button) return;
                const input = listEl.querySelector(`.sn-formula-arg-input[data-arg-index="${button.dataset.argIndex}"]`);
                if (!input) return;
                setActiveInput(input);
                input.focus();
            });

            addButton?.addEventListener('click', () => {
                const index = listEl.querySelectorAll('.sn-formula-arg-input').length;
                const definition = {
                    label: SN.t('action.formula.content.arguments.argument', { index: index + 1 }),
                    optional: true
                };
                listEl.insertAdjacentHTML('beforeend', buildArgumentRowHTML(definition, index, SN));
                const input = listEl.querySelector(`.sn-formula-arg-input[data-arg-index="${index}"]`);
                setActiveInput(input);
                input?.focus();
                input?.scrollIntoView({ block: 'nearest' });
            });
            argsEl?.addEventListener('keydown', (event) => {
                event.stopPropagation();
                if (event.isComposing) return;
                if (event.key === 'Escape') {
                    event.preventDefault();
                    modalEl?.querySelector('.sn-modal-btn-cancel')?.click();
                } else if (event.key === 'Enter' && event.target.closest('.sn-formula-arg-input')) {
                    event.preventDefault();
                    modalEl?.querySelector('.sn-modal-btn-confirm')?.click();
                }
            });

            if (activeInput) {
                setActiveInput(activeInput);
                if (pendingReference) activeInput.value = pendingReference;
                activeInput.focus();
            } else {
                modalEl?.querySelector('.sn-modal-btn-confirm')?.focus();
            }
            releaseFormulaEditSuspension(SN);
        },
        onConfirm: (bodyEl) => {
            const values = Array.from(bodyEl.querySelectorAll('.sn-formula-arg-input'))
                .map(input => input.value.trim());
            while (values.length && values[values.length - 1] === '') values.pop();
            const call = `${name}(${values.join(',')})`;
            const draft = context.draft.trim();
            let nextFormula = `=${call}`;
            if (context.wasEditing && draft.startsWith('=')) {
                const start = Math.min(Math.max(context.selectionStart, 1), draft.length);
                const end = Math.min(Math.max(context.selectionEnd, start), draft.length);
                nextFormula = draft.slice(0, start) + call + draft.slice(end);
            }
            const targetCell = context.targetSheet.getCell(context.targetCell.r, context.targetCell.c);
            if (targetCell.dataValidation?.showErrorMessage) {
                let candidateValue = SN.Formula.calcFormula(nextFormula.slice(1), targetCell);
                if (Array.isArray(candidateValue) && Array.isArray(candidateValue[0])) {
                    candidateValue = candidateValue[0]?.[0] ?? '';
                } else if (candidateValue instanceof Error) {
                    candidateValue = candidateValue.message;
                }
                if (!targetCell.validateData(candidateValue)) {
                    const validationMessage = targetCell.dataValidation.error
                        || SN.t('dataValidation.errors.inputRuleMismatch');
                    alert(SN.t(validationMessage));
                    return false;
                }
            }
            targetCell.editVal = nextFormula;
            if (targetCell.editVal !== nextFormula) return false;
            if (SN.DependencyGraph && SN.calcMode !== 'manual') {
                SN.DependencyGraph.recalculateVolatile();
            }
            FormulaCatalog.addRecent(name);
            committed = true;
            stopRangePicker();
            restoreFormulaTarget(SN, context, false, true);
            return true;
        },
        onClosed: () => {
            this._formulaArgumentDialogOpen = false;
            releaseFormulaEditSuspension(SN);
            stopRangePicker();
            if (!committed) restoreFormulaTarget(SN, context, true);
        }
    });
}

export function openFormulaDialog(categoryKey = 'all') {
    if (this._formulaArgumentDialogOpen) {
        focusOpenArgumentDialog(this.SN);
        return;
    }
    const categories = FormulaCatalog.getDialogCategories();
    const validKeys = new Set(categories.map(item => item.key));
    const initialKey = validKeys.has(categoryKey) ? categoryKey : 'all';
    const insertContext = captureFormulaInsertContext(this.SN);
    let selectedName = '';
    suspendFormulaEdit(this.SN);

    this.SN.Utils.modal({
        titleKey: 'action.formula.modal.m001.title', title: 'Insert Function',
        maxWidth: '420px',
        confirmKey: 'action.formula.modal.m001.confirm', confirmText: 'Insert',
        cancelKey: 'action.formula.modal.m001.cancel', cancelText: 'Cancel',
        content: buildDialogHTML(categories, initialKey, this.SN),
        onOpen: (bodyEl) => {
            suspendFormulaEdit(this.SN);
            bodyEl.classList.add('sn-modal-body-formula');
            bodyEl.querySelector('.sn-formula-dialog')?.addEventListener('keydown', event => event.stopPropagation());
            const categorySelect = bodyEl.querySelector('.sn-formula-category-select');
            const searchInput = bodyEl.querySelector('.sn-formula-search-input');
            const listEl = bodyEl.querySelector('.sn-formula-list');
            const overlay = bodyEl.closest('.sn-modal-overlay');
            const confirmBtn = overlay?.querySelector('.sn-modal-btn-confirm');

            let activeCategory = initialKey;
            let currentSelection = '';

            const render = () => {
                const list = FormulaCatalog.search(searchInput.value, activeCategory);
                const hasSelection = currentSelection && list.some(item => item.name === currentSelection);
                const activeName = hasSelection ? currentSelection : (list[0]?.name || '');
                renderFormulaList(listEl, list, activeName, this.SN);
                currentSelection = activeName;
                renderFormulaDetail(bodyEl, FormulaCatalog.getFormula(currentSelection), this.SN);
            };

            categorySelect.addEventListener('change', () => {
                activeCategory = categorySelect.value;
                currentSelection = '';
                render();
            });

            listEl.addEventListener('click', (e) => {
                const item = e.target.closest('.sn-formula-item');
                if (!item) return;
                listEl.querySelectorAll('.sn-formula-item').forEach(el => el.classList.remove('active'));
                item.classList.add('active');
                currentSelection = item.dataset.name || '';
                renderFormulaDetail(bodyEl, FormulaCatalog.getFormula(currentSelection), this.SN);
            });

            listEl.addEventListener('dblclick', (e) => {
                const item = e.target.closest('.sn-formula-item');
                if (!item || !confirmBtn) return;
                currentSelection = item.dataset.name || '';
                confirmBtn.click();
            });

            searchInput.addEventListener('input', () => {
                currentSelection = '';
                render();
            });

            render();
            searchInput.focus();
            releaseFormulaEditSuspension(this.SN);
        },
        onConfirm: (bodyEl) => {
            const selected = bodyEl?.querySelector('.sn-formula-item.active');
            if (!selected) {
                this.SN.Utils.toast(this.SN.t('action.formula.toast.pleaseSelectAFunction'));
                return false;
            }
            selectedName = selected.dataset.name || '';
            return true;
        },
        onClosed: (confirmed) => {
            releaseFormulaEditSuspension(this.SN);
            if (confirmed && selectedName) {
                this.insertFunction(selectedName, insertContext);
            } else {
                restoreFormulaTarget(this.SN, insertContext, true);
            }
        }
    });
}

export function insertFunction(name, insertContext = null) {
    if (this._formulaArgumentDialogOpen) {
        focusOpenArgumentDialog(this.SN);
        return;
    }
    if (!name) {
        this.openFormulaDialog('all');
        return;
    }

    const formula = FormulaCatalog.getFormula(name);
    const finalName = formula?.name || String(name).toUpperCase();

    const input = this.SN.Canvas?.formulaBar;
    if (!input) return;

    if (formula) {
        const context = insertContext || captureFormulaInsertContext(this.SN);
        suspendFormulaEdit(this.SN);
        openFunctionArguments.call(this, finalName, context);
        return;
    }

    let value = input.value || '';
    if (!value.trim().startsWith('=')) {
        value = value.trim() ? '=' + value : '=';
    }

    const insertText = `${finalName}(`;
    input.value = value + insertText;
    const caretPos = input.value.length;

    if (this.SN.Canvas?.input) this.SN.Canvas.input.innerText = input.value;
    input.focus();
    input.setSelectionRange?.(caretPos, caretPos);
}
