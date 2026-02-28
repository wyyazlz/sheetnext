import { FormulaCatalog } from '../core/Formula/FormulaCatalog.js';

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

export function openFormulaDialog(categoryKey = 'all') {
    const categories = FormulaCatalog.getDialogCategories();
    const validKeys = new Set(categories.map(item => item.key));
    const initialKey = validKeys.has(categoryKey) ? categoryKey : 'all';

    this.SN.Utils.modal({
        titleKey: 'action.formula.modal.m001.title', title: 'Insert Function',
        maxWidth: '420px',
        confirmKey: 'action.formula.modal.m001.confirm', confirmText: 'Insert',
        cancelKey: 'action.formula.modal.m001.cancel', cancelText: 'Cancel',
        content: buildDialogHTML(categories, initialKey, this.SN),
        onOpen: (bodyEl) => {
            bodyEl.classList.add('sn-modal-body-formula');
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
        },
        onConfirm: (bodyEl) => {
            const selected = bodyEl?.querySelector('.sn-formula-item.active');
            if (!selected) {
                this.SN.Utils.toast(this.SN.t('action.formula.toast.pleaseSelectAFunction'));
                return false;
            }
            this.insertFunction(selected.dataset.name);
            return true;
        }
    });
}

export function insertFunction(name) {
    if (!name) {
        this.openFormulaDialog('all');
        return;
    }

    const formula = FormulaCatalog.getFormula(name);
    const finalName = formula?.name || String(name).toUpperCase();
    FormulaCatalog.addRecent(finalName);

    const input = this.SN.Canvas?.formulaBar;
    if (!input) return;

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
