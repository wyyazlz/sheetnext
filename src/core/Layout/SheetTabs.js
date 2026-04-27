function escapeHtml(value) {
    const div = document.createElement('div');
    div.textContent = value == null ? '' : String(value);
    return div.innerHTML;
}

function getSheetFromTab(SN, tab) {
    const index = Number(tab?.dataset?.sheetIndex);
    if (!Number.isInteger(index)) return null;
    return SN.sheets[index] || null;
}

function getSheetByMenu(menu, SN) {
    const index = Number(menu?.dataset?.sheetIndex);
    if (!Number.isInteger(index)) return null;
    return SN.sheets[index] || null;
}

function getSheetDialogContent(label, selectHTML) {
    return `
        <div class="sn-sheet-dialog">
            <label>${escapeHtml(label)}</label>
            <select data-field="sheet-target">${selectHTML}</select>
        </div>
    `;
}

/** Bind worksheet tab events once. */
export function _initSheetTabEvents() {
    const list = this.SN.containerDom.querySelector('.sn-sheet-list');
    if (!list || list._snSheetTabEventsBound) return;
    list._snSheetTabEventsBound = true;

    const findTab = (event) => {
        const tab = event.target?.closest?.('.sn-sheet-item');
        return tab && list.contains(tab) ? tab : null;
    };

    list.addEventListener('mousedown', (event) => {
        if (event.button !== 0 || event.target?.closest?.('.sn-sheet-rename-input')) return;
        const tab = findTab(event);
        if (tab) this._activateSheetTab(tab);
    });

    list.addEventListener('touchstart', (event) => {
        if (event.target?.closest?.('.sn-sheet-rename-input')) return;
        const tab = findTab(event);
        if (tab) this._activateSheetTab(tab);
    }, { passive: true });

    list.addEventListener('dblclick', (event) => {
        const tab = findTab(event);
        if (!tab) return;
        event.preventDefault();
        this._beginSheetTabRename(tab);
    });

    list.addEventListener('contextmenu', (event) => {
        const tab = findTab(event);
        if (!tab) return;
        this._showSheetTabContextMenu(event, tab);
    });

    this.SN.containerDom.addEventListener('mousedown', (event) => {
        const menu = this.SN.Canvas?.lazyDOM?.cache?.get('sheet-context-menu');
        if (!menu || menu.style.display === 'none') return;
        if (menu.contains(event.target) || event.target?.closest?.('.sn-sheet-item')) return;
        menu.style.display = 'none';
    });
}

export function _activateSheetTab(tab) {
    const sheet = getSheetFromTab(this.SN, tab);
    if (!sheet || sheet.hidden) return;
    if (sheet !== this.SN.activeSheet) {
        this.SN.activeSheet = sheet;
    } else {
        tab.classList.add('sheet-sel');
    }
}

export function _showSheetTabContextMenu(event, tab) {
    event.preventDefault();
    event.stopPropagation();

    const sheet = getSheetFromTab(this.SN, tab);
    if (!sheet) return;
    if (sheet !== this.SN.activeSheet && !sheet.hidden) this.SN.activeSheet = sheet;

    const menu = this.SN.Canvas?.lazyDOM?.sheetContextMenu;
    if (!menu) return;
    menu.dataset.sheetIndex = String(this.SN.sheets.indexOf(sheet));
    this.SN.Canvas.lazyDOM._syncSheetContextMenu(sheet);

    if (!menu._snSheetMenuEventsBound) {
        menu._snSheetMenuEventsBound = true;
        menu.addEventListener('click', (clickEvent) => this._handleSheetTabMenuClick(clickEvent));
    }

    this._positionSheetTabContextMenu(menu, event.clientX, event.clientY);
}

export function _handleSheetTabMenuClick(event) {
    const item = event.target?.closest?.('.sn-rc-item[data-sheet-action]');
    if (!item || item.classList.contains('sn-rc-item-disabled')) return;

    const menu = item.closest('.sn-sheet-rc');
    const sheet = getSheetByMenu(menu, this.SN);
    if (!sheet) return;

    menu.style.display = 'none';
    const action = item.dataset.sheetAction;
    const sheetIndex = this.SN.sheets.indexOf(sheet);

    if (action === 'insert') {
        const newSheet = this.SN.addSheet();
        if (newSheet) {
            this.SN.moveSheet(newSheet, sheetIndex);
            this.SN.activeSheet = newSheet;
        }
        return;
    }

    if (action === 'delete') {
        this.SN.Utils.modal({
            titleKey: 'workbook.sheetTab.deleteTitle',
            title: this.SN.t('workbook.sheetTab.deleteTitle'),
            content: escapeHtml(this.SN.t('workbook.sheetTab.deleteMessage', { name: sheet.name })),
            confirmKey: 'workbook.sheetTab.deleteConfirm',
            confirmText: this.SN.t('workbook.sheetTab.deleteConfirm'),
            cancelKey: 'common.cancel',
            cancelText: this.SN.t('common.cancel'),
            maxWidth: '360px',
            onConfirm: () => {
                this.SN.delSheet(sheet.name);
                return true;
            }
        });
        return;
    }

    if (action === 'rename') {
        this._beginSheetTabRename(sheet);
        return;
    }

    if (action === 'hide') {
        sheet.hidden = true;
        this.SN._updListDom();
        return;
    }

    if (action === 'unhide') {
        this._openUnhideSheetDialog();
        return;
    }

    if (action === 'move') {
        this._openMoveSheetDialog(sheet);
    }
}

export function _beginSheetTabRename(target) {
    const sheet = target?.classList?.contains?.('sn-sheet-item')
        ? getSheetFromTab(this.SN, target)
        : target;
    if (!sheet) return;

    const index = this.SN.sheets.indexOf(sheet);
    const tab = this.SN.containerDom.querySelector(`.sn-sheet-item[data-sheet-index="${index}"]`);
    if (!tab || tab.querySelector('.sn-sheet-rename-input')) return;

    const oldName = sheet.name;
    tab.style.width = `${Math.max(72, tab.offsetWidth)}px`;
    tab.classList.add('sn-sheet-renaming');
    tab.textContent = '';

    const input = document.createElement('input');
    input.className = 'sn-sheet-rename-input';
    input.value = oldName;
    input.setAttribute('aria-label', this.SN.t('workbook.sheetTab.renameLabel'));
    tab.appendChild(input);

    let done = false;
    const finish = (commit) => {
        if (done) return;
        done = true;
        if (commit && input.value !== oldName) sheet.name = input.value;
        if (sheet.name === oldName) this.SN._updListDom();
    };

    input.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
            event.preventDefault();
            finish(true);
        } else if (event.key === 'Escape') {
            event.preventDefault();
            finish(false);
        }
    });
    input.addEventListener('blur', () => finish(true));

    requestAnimationFrame(() => {
        input.focus();
        input.select();
    });
}

export function _openUnhideSheetDialog() {
    const hiddenSheets = this.SN.sheets.filter(sheet => sheet.hidden);
    if (!hiddenSheets.length) {
        this.SN.Utils.toast(this.SN.t('workbook.sheetTab.noHiddenSheets'));
        return;
    }

    const options = hiddenSheets.map(sheet => {
        const index = this.SN.sheets.indexOf(sheet);
        return `<option value="${index}">${escapeHtml(sheet.name)}</option>`;
    }).join('');

    this.SN.Utils.modal({
        titleKey: 'workbook.sheetTab.unhideTitle',
        title: this.SN.t('workbook.sheetTab.unhideTitle'),
        content: getSheetDialogContent(this.SN.t('workbook.sheetTab.unhideLabel'), options),
        confirmKey: 'common.ok',
        confirmText: this.SN.t('common.ok'),
        cancelKey: 'common.cancel',
        cancelText: this.SN.t('common.cancel'),
        maxWidth: '360px',
        onConfirm: (bodyEl) => {
            const index = Number(bodyEl.querySelector('[data-field="sheet-target"]')?.value);
            const sheet = this.SN.sheets[index];
            if (!sheet) return false;
            sheet.hidden = false;
            this.SN._updListDom();
            return true;
        }
    });
}

export function _openMoveSheetDialog(sheet) {
    const fromIndex = this.SN.sheets.indexOf(sheet);
    if (fromIndex === -1) return;

    const options = this.SN.sheets.map((item, index) => {
        const selected = index === fromIndex ? ' selected' : '';
        const label = this.SN.t('workbook.sheetTab.moveBeforeSheet', { name: item.name });
        return `<option value="${index}"${selected}>${escapeHtml(label)}</option>`;
    }).join('') + `<option value="${this.SN.sheets.length}">${escapeHtml(this.SN.t('workbook.sheetTab.moveToEnd'))}</option>`;

    this.SN.Utils.modal({
        titleKey: 'workbook.sheetTab.moveTitle',
        title: this.SN.t('workbook.sheetTab.moveTitle'),
        content: getSheetDialogContent(this.SN.t('workbook.sheetTab.moveBeforeLabel'), options),
        confirmKey: 'workbook.sheetTab.moveConfirm',
        confirmText: this.SN.t('workbook.sheetTab.moveConfirm'),
        cancelKey: 'common.cancel',
        cancelText: this.SN.t('common.cancel'),
        maxWidth: '380px',
        onConfirm: (bodyEl) => {
            const rawIndex = Number(bodyEl.querySelector('[data-field="sheet-target"]')?.value);
            const currentIndex = this.SN.sheets.indexOf(sheet);
            if (!Number.isInteger(rawIndex) || currentIndex === -1) return false;

            const isMoveToEnd = rawIndex >= this.SN.sheets.length;
            let targetIndex = isMoveToEnd ? this.SN.sheets.length - 1 : rawIndex;
            if (!isMoveToEnd && rawIndex > currentIndex) targetIndex -= 1;
            this.SN.moveSheet(sheet, targetIndex);
            this.SN.activeSheet = sheet;
            return true;
        }
    });
}

export function _positionSheetTabContextMenu(menu, clientX, clientY) {
    const host = this.SN.containerDom.querySelector('.sn-main') || this.SN.containerDom;
    const rect = host.getBoundingClientRect();
    menu.style.display = 'block';

    const menuRect = menu.getBoundingClientRect();
    const gap = 6;
    const maxLeft = Math.max(gap, rect.width - menuRect.width - gap);
    const maxTop = Math.max(gap, rect.height - menuRect.height - gap);
    const left = Math.max(gap, Math.min(clientX - rect.left + 2, maxLeft));
    const top = Math.max(gap, Math.min(clientY - rect.top + 2, maxTop));

    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
}
