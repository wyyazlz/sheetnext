function _renderScopeLabel(scope, t) {
    return scope === 'workbook'
        ? t('action.nameManager.content.scope.workbook')
        : scope;
}

function _buildNameManagerRows(names, selectedName, t) {
    if (!names.length) {
        return `<tr><td colspan="5" class="sn-nm-empty">${t('action.nameManager.content.empty')}</td></tr>`;
    }
    return names.map((item, idx) => {
        const comment = item.comment ?? item.Comment ?? '';
        const scope = item.scope || 'workbook';
        return `
            <tr class="sn-nm-row${item.name === selectedName || (!selectedName && idx === 0) ? ' active' : ''}" data-name="${item.name}">
                <td class="sn-nm-cell-name">${item.name}</td>
                <td class="sn-nm-cell-value" title="${item.value || ''}">${item.value || ''}</td>
                <td class="sn-nm-cell-ref" title="${item.ref || ''}">${item.ref || ''}</td>
                <td class="sn-nm-cell-scope">${_renderScopeLabel(scope, t)}</td>
                <td class="sn-nm-cell-comment" title="${comment}">${comment}</td>
            </tr>
        `;
    }).join('');
}

function _buildNameManagerHTML(names, filterScope, t) {
    const scopeOptions = `
        <option value="all"${filterScope === 'all' ? ' selected' : ''}>${t('action.nameManager.content.scope.all')}</option>
        <option value="workbook"${filterScope === 'workbook' ? ' selected' : ''}>${t('action.nameManager.content.scope.workbook')}</option>
    `;
    const rows = _buildNameManagerRows(names, names[0]?.name || '', t);
    return `
        <div class="sn-name-manager">
            <div class="sn-nm-toolbar">
                <button class="sn-nm-btn sn-nm-btn-new">${t('action.nameManager.content.button.new')}</button>
                <button class="sn-nm-btn sn-nm-btn-edit" disabled>${t('action.nameManager.content.button.edit')}</button>
                <button class="sn-nm-btn sn-nm-btn-delete" disabled>${t('action.nameManager.content.button.delete')}</button>
                <div class="sn-nm-filter">
                    <label>${t('action.nameManager.content.filterLabel')}</label>
                    <select class="sn-nm-filter-select">${scopeOptions}</select>
                </div>
            </div>
            <div class="sn-nm-table-wrap">
                <table class="sn-nm-table">
                    <thead>
                        <tr>
                            <th style="width:120px">${t('action.nameManager.content.column.name')}</th>
                            <th style="width:100px">${t('action.nameManager.content.column.value')}</th>
                            <th style="width:140px">${t('action.nameManager.content.column.ref')}</th>
                            <th style="width:80px">${t('action.nameManager.content.column.scope')}</th>
                            <th>${t('action.nameManager.content.column.comment')}</th>
                        </tr>
                    </thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>
            <div class="sn-nm-footer">
                <label>${t('action.nameManager.content.refLabel')}</label>
                <input type="text" class="sn-nm-ref-input" readonly>
            </div>
        </div>
    `;
}

function _buildDefineNameHTML(data, sheets, t) {
    const { name = '', scope = 'workbook', comment = '', ref = '' } = data;
    const scopeOptions = [
        `<option value="workbook"${scope === 'workbook' ? ' selected' : ''}>${t('action.nameManager.content.scope.workbook')}</option>`,
        ...sheets.map((sheet) => `<option value="${sheet.name}"${scope === sheet.name ? ' selected' : ''}>${sheet.name}</option>`)
    ].join('');

    return `
        <div class="sn-define-name">
            <div class="sn-dn-row">
                <label>${t('action.nameManager.content.form.nameLabel')}</label>
                <input type="text" class="sn-dn-name" value="${name}" placeholder="${t('action.nameManager.content.form.namePlaceholder')}">
            </div>
            <div class="sn-dn-row">
                <label>${t('action.nameManager.content.form.scopeLabel')}</label>
                <select class="sn-dn-scope">${scopeOptions}</select>
            </div>
            <div class="sn-dn-row">
                <label>${t('action.nameManager.content.form.commentLabel')}</label>
                <input type="text" class="sn-dn-comment" value="${comment}" placeholder="${t('action.nameManager.content.form.commentPlaceholder')}">
            </div>
            <div class="sn-dn-row">
                <label>${t('action.nameManager.content.form.refLabel')}</label>
                <input type="text" class="sn-dn-ref" value="${ref}" placeholder="${t('action.nameManager.content.form.refPlaceholder')}">
            </div>
        </div>
    `;
}

function _formatAbsoluteRef(sheetName, rangeStr) {
    const parts = rangeStr.split(':');
    const _formatCell = (cell) => {
        const match = cell.match(/^([A-Z]+)(\d+)$/i);
        if (!match) return cell;
        return `$${match[1].toUpperCase()}$${match[2]}`;
    };
    const formattedRange = parts.map(_formatCell).join(':');
    return `${sheetName}!${formattedRange}`;
}

function _getDefinedNamesList() {
    const map = this.SN._definedNames || {};
    return Object.entries(map).map(([name, ref]) => ({
        name,
        ref,
        value: '',
        scope: 'workbook',
        comment: ''
    }));
}

function _filterNamesByScope(names, scope) {
    if (scope === 'all') return names;
    return names.filter((item) => item.scope === scope);
}

export function nameManager() {
    const t = (key, params) => this.SN.t(key, params);
    const names = _getDefinedNamesList.call(this);

    this.SN.Utils.modal({
        titleKey: 'action.nameManager.modal.m001.title',
        title: t('action.nameManager.modal.m001.title'),
        maxWidth: '640px',
        confirmKey: 'action.nameManager.modal.m001.confirm',
        confirmText: t('action.nameManager.modal.m001.confirm'),
        content: _buildNameManagerHTML(names, 'all', t),
        onOpen: (bodyEl) => {
            bodyEl.classList.add('sn-modal-body-nm');
            const tbody = bodyEl.querySelector('.sn-nm-table tbody');
            const refInput = bodyEl.querySelector('.sn-nm-ref-input');
            const btnNew = bodyEl.querySelector('.sn-nm-btn-new');
            const btnEdit = bodyEl.querySelector('.sn-nm-btn-edit');
            const btnDelete = bodyEl.querySelector('.sn-nm-btn-delete');
            const filterSelect = bodyEl.querySelector('.sn-nm-filter-select');

            let selectedName = names[0]?.name || '';
            let currentFilter = filterSelect?.value || 'all';

            const _updateSelection = () => {
                const item = names.find((nameItem) => nameItem.name === selectedName);
                refInput.value = item?.ref || '';
                btnEdit.disabled = !selectedName;
                btnDelete.disabled = !selectedName;
            };

            const _refreshList = () => {
                const list = _getDefinedNamesList.call(this);
                names.length = 0;
                names.push(...list);
                const visibleNames = _filterNamesByScope(list, currentFilter);
                if (!visibleNames.some((item) => item.name === selectedName)) {
                    selectedName = visibleNames[0]?.name || '';
                }
                tbody.innerHTML = _buildNameManagerRows(visibleNames, selectedName, t);
                _updateSelection();
            };

            tbody.addEventListener('click', (event) => {
                const row = event.target.closest('.sn-nm-row');
                if (!row) return;
                tbody.querySelectorAll('.sn-nm-row').forEach((item) => item.classList.remove('active'));
                row.classList.add('active');
                selectedName = row.dataset.name;
                _updateSelection();
            });

            btnNew.addEventListener('click', () => {
                this.defineName(() => _refreshList());
            });

            btnEdit.addEventListener('click', () => {
                if (!selectedName) return;
                const item = names.find((nameItem) => nameItem.name === selectedName);
                if (item) this.defineName(() => _refreshList(), item);
            });

            btnDelete.addEventListener('click', () => {
                if (!selectedName) return;
                const nameToDelete = selectedName;
                const refToDelete = this.SN._definedNames[nameToDelete];
                this.SN.Utils.modal({
                    titleKey: 'action.nameManager.content.deleteConfirmTitle',
                    title: t('action.nameManager.content.deleteConfirmTitle'),
                    content: t('action.nameManager.content.deleteConfirmBody', { name: nameToDelete }),
                    confirmKey: 'action.nameManager.content.deleteConfirmButton',
                    confirmText: t('action.nameManager.content.deleteConfirmButton'),
                    cancelKey: 'action.nameManager.modal.m001.cancel',
                    cancelText: t('action.nameManager.modal.m001.cancel'),
                    onConfirm: () => {
                        delete this.SN._definedNames[nameToDelete];
                        this.SN.UndoRedo.add({
                            undo: () => { this.SN._definedNames[nameToDelete] = refToDelete; },
                            redo: () => { delete this.SN._definedNames[nameToDelete]; }
                        });
                        selectedName = '';
                        _refreshList();
                        return true;
                    }
                });
            });

            filterSelect.addEventListener('change', () => {
                currentFilter = filterSelect.value || 'all';
                _refreshList();
            });

            _updateSelection();
        },
        onConfirm: () => true
    });
}

export function defineName(callback, editData = null) {
    const t = (key, params) => this.SN.t(key, params);
    const isEdit = !!editData;
    const sheet = this.SN.activeSheet;
    const area = sheet.activeAreas[0] || { s: sheet.activeCell, e: sheet.activeCell };
    const rangeStr = this.SN.Utils.rangeNumToStr(area);
    const defaultRef = _formatAbsoluteRef(sheet.name, rangeStr);
    const data = editData || { ref: defaultRef };

    const modalTitleKey = isEdit
        ? 'action.nameManager.modal.m002.titleEdit'
        : 'action.nameManager.modal.m002.titleCreate';

    this.SN.Utils.modal({
        titleKey: modalTitleKey,
        title: t(modalTitleKey),
        maxWidth: '400px',
        content: _buildDefineNameHTML(data, this.SN.sheets, t),
        confirmKey: 'action.nameManager.modal.m002.confirm',
        confirmText: t('action.nameManager.modal.m002.confirm'),
        cancelKey: 'action.nameManager.modal.m002.cancel',
        cancelText: t('action.nameManager.modal.m002.cancel'),
        onOpen: (bodyEl) => {
            bodyEl.classList.add('sn-modal-body-dn');
            const nameInput = bodyEl.querySelector('.sn-dn-name');

            if (isEdit) {
                nameInput.readOnly = true;
                nameInput.style.background = '#f5f5f5';
            } else {
                nameInput.focus();
            }
        },
        onConfirm: (bodyEl) => {
            const nameInput = bodyEl.querySelector('.sn-dn-name');
            const refInput = bodyEl.querySelector('.sn-dn-ref');
            const name = nameInput.value.trim();
            const ref = refInput.value.trim();

            if (!name) {
                this.SN.Utils.toast(t('action.nameManager.toast.msg001'));
                nameInput.focus();
                return false;
            }

            if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) {
                this.SN.Utils.toast(t('action.nameManager.toast.msg002'));
                nameInput.focus();
                return false;
            }

            if (!ref) {
                this.SN.Utils.toast(t('action.nameManager.toast.msg003'));
                refInput.focus();
                return false;
            }

            if (!isEdit && this.SN._definedNames[name]) {
                this.SN.Utils.toast(t('action.nameManager.toast.msg004'));
                nameInput.focus();
                return false;
            }

            const oldRef = isEdit ? editData.ref : null;
            this.SN._definedNames[name] = ref;

            if (isEdit) {
                this.SN.UndoRedo.add({
                    undo: () => { this.SN._definedNames[name] = oldRef; },
                    redo: () => { this.SN._definedNames[name] = ref; }
                });
            } else {
                this.SN.UndoRedo.add({
                    undo: () => { delete this.SN._definedNames[name]; },
                    redo: () => { this.SN._definedNames[name] = ref; }
                });
            }

            callback?.();
            return true;
        }
    });
}

export function useInFormula() {
    const t = (key, params) => this.SN.t(key, params);
    const names = _getDefinedNamesList.call(this);
    if (!names.length) {
        this.SN.Utils.toast(t('action.nameManager.toast.msg005'));
        return;
    }

    this.SN.Utils.modal({
        titleKey: 'action.nameManager.modal.m003.title',
        title: t('action.nameManager.modal.m003.title'),
        maxWidth: '320px',
        content: `
            <div class="sn-use-formula">
                <div class="sn-uf-list">
                    ${names.map((nameItem, idx) => `
                        <div class="sn-uf-item${idx === 0 ? ' active' : ''}" data-name="${nameItem.name}">
                            <span class="sn-uf-name">${nameItem.name}</span>
                            <span class="sn-uf-ref">${nameItem.ref}</span>
                        </div>
                    `).join('')}
                </div>
            </div>
        `,
        confirmKey: 'action.nameManager.modal.m003.confirm',
        confirmText: t('action.nameManager.modal.m003.confirm'),
        cancelKey: 'action.nameManager.modal.m003.cancel',
        cancelText: t('action.nameManager.modal.m003.cancel'),
        onOpen: (bodyEl) => {
            const list = bodyEl.querySelector('.sn-uf-list');
            list.addEventListener('click', (event) => {
                const item = event.target.closest('.sn-uf-item');
                if (!item) return;
                list.querySelectorAll('.sn-uf-item').forEach((node) => node.classList.remove('active'));
                item.classList.add('active');
            });
            list.addEventListener('dblclick', (event) => {
                const item = event.target.closest('.sn-uf-item');
                if (!item) return;
                const overlay = bodyEl.closest('.sn-modal-overlay');
                overlay?.querySelector('.sn-modal-btn-confirm')?.click();
            });
        },
        onConfirm: (bodyEl) => {
            const selected = bodyEl.querySelector('.sn-uf-item.active');
            if (!selected) return false;
            this.insertFunction(selected.dataset.name);
            return true;
        }
    });
}
