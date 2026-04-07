/**
 * Filter panel interaction module.
 * Creates the panel on demand and handles user interaction.
 */

import getSvg from '../../assets/mainSvgs.js';

export default class FilterPanel {
    /** @param {import('../Canvas/Canvas.js').default} canvas */
    constructor(canvas) {
        /** @type {import('../Canvas/Canvas.js').default} */
        this.canvas = canvas;
        this._SN = canvas.SN;
        this._panel = null;
        this._currentColIndex = null;
        this._currentScopeId = null;
        this._allValues = [];
        this._selectedValues = new Set();
        this._boundClickOutside = this._onClickOutside.bind(this);
        this._customContext = null;
        this._showCount = true;
        this._filterMode = 'list';
        this._dateTree = null;
        this._collapsedDateNodes = new Set();
        this._searchKeyword = '';
    }

    /**
     * Get active AutoFilter instance.
     * @type {import('./AutoFilter.js').default|null}
     */
    get autoFilter() {
        return this._SN.activeSheet?.AutoFilter;
    }

    _t(key, params = {}) {
        if (!key || typeof key !== 'string') return '';
        return this._SN.t(key, params);
    }

    _createPanel() {
        if (this._panel) return this._panel;

        const panel = document.createElement('div');
        panel.className = 'sn-filter-panel';
        panel.innerHTML = `
            <div class="sn-filter-header">
                <span class="sn-filter-title">${this._t('autoFilter.panel.title')}</span>
                <span class="sn-filter-close">×</span>
            </div>
            <div class="sn-filter-field">
                <label>${this._t('autoFilter.panel.field')}</label>
                <select class="sn-filter-field-select"></select>
            </div>
            <div class="sn-filter-sort">
                <div class="sn-filter-sort-item" data-sort="asc">
                    ${getSvg('filter_sortAsc')}
                    <span>${this._t('autoFilter.panel.sortAsc')}</span>
                </div>
                <div class="sn-filter-sort-item" data-sort="desc">
                    ${getSvg('filter_sortDesc')}
                    <span>${this._t('autoFilter.panel.sortDesc')}</span>
                </div>
            </div>
            <div class="sn-filter-search">
                <input type="text" placeholder="${this._t('autoFilter.panel.searchPlaceholder')}" class="sn-filter-search-input">
            </div>
            <div class="sn-filter-actions">
                <span class="sn-filter-select-all">${this._t('autoFilter.panel.selectAll')}</span>
                <span class="sn-filter-clear-all">${this._t('autoFilter.panel.clear')}</span>
            </div>
            <div class="sn-filter-list"></div>
            <div class="sn-filter-footer">
                <button class="sn-filter-cancel">${this._t('common.cancel')}</button>
                <button class="sn-filter-apply primary">${this._t('common.ok')}</button>
            </div>
        `;

        document.body.appendChild(panel);
        this._panel = panel;
        this._bindEvents();
        return panel;
    }

    _bindEvents() {
        const panel = this._panel;
        if (!panel) return;

        panel.querySelector('.sn-filter-close')?.addEventListener('click', () => this.hide());

        const searchInput = panel.querySelector('.sn-filter-search-input');
        searchInput?.addEventListener('input', (e) => this._onSearch(e.target.value));

        panel.querySelector('.sn-filter-select-all')?.addEventListener('click', () => this._selectAll());
        panel.querySelector('.sn-filter-clear-all')?.addEventListener('click', () => this._clearAll());

        panel.querySelectorAll('.sn-filter-sort-item').forEach(item => {
            item.addEventListener('click', () => {
                this._doSort(item.dataset.sort);
            });
        });

        panel.querySelector('.sn-filter-apply')?.addEventListener('click', () => this._apply());
        panel.querySelector('.sn-filter-cancel')?.addEventListener('click', () => this.hide());
        panel.addEventListener('mousedown', (e) => e.stopPropagation());

        const fieldSelect = panel.querySelector('.sn-filter-field-select');
        fieldSelect?.addEventListener('change', () => {
            if (this._customContext?.type === 'pivot') {
                this._setPivotField(Number(fieldSelect.value));
            }
        });
    }

    _onClickOutside(e) {
        if (!this._panel) return;
        const isFilterIcon = e.target.closest && e.target.closest('.sn-filter-icon');
        if (!isFilterIcon && !this._panel.contains(e.target)) {
            this.hide();
        }
    }

    /** @param {number} colIndex @param {number} x @param {number} y @param {string|null} scopeId */
    show(colIndex, x, y, scopeId = null) {
        const autoFilter = this.autoFilter;
        if (!autoFilter?.isScopeEnabled(scopeId)) return;

        document.removeEventListener('mousedown', this._boundClickOutside);
        this._createPanel();

        this._currentColIndex = colIndex;
        this._currentScopeId = scopeId;
        this._customContext = null;
        this._showCount = true;
        this._filterMode = 'list';
        this._dateTree = null;
        this._collapsedDateNodes = new Set();
        this._searchKeyword = '';
        this._toggleFieldSelect(false);

        this._allValues = autoFilter.getColumnValues(colIndex, this._currentScopeId);
        this._dateTree = this._buildDateTree(this._allValues);
        this._initDateTreeCollapsedState();
        this._filterMode = this._dateTree ? 'date' : 'list';

        const currentFilter = autoFilter.getColumnFilter(colIndex, this._currentScopeId);
        if (currentFilter?.type === 'values') {
            this._selectedValues = new Set(currentFilter.values.map(v => String(v ?? '')));
        } else {
            this._selectedValues = new Set(this._allValues.map(v => v.filterKey ?? String(v.value ?? '')));
        }

        this._renderCurrentView();

        const searchInput = this._panel.querySelector('.sn-filter-search-input');
        if (searchInput) searchInput.value = '';

        const title = this._panel.querySelector('.sn-filter-title');
        if (title) {
            const colName = this._SN.Utils.numToChar(colIndex);
            title.textContent = this._t('autoFilter.panel.titleColumn', { column: colName });
        }

        const currentSort = autoFilter.getSortOrder(colIndex, this._currentScopeId);
        this._panel.querySelectorAll('.sn-filter-sort-item').forEach(item => {
            item.classList.toggle('active', item.dataset.sort === currentSort);
        });

        this._positionPanel(x, y);
    }

    /** @param {Object} pivotTable @param {Object} filterInfo @param {number} x @param {number} y */
    showPivotFilter(pivotTable, filterInfo, x, y) {
        const fields = Array.isArray(filterInfo?.fields) ? filterInfo.fields : [];
        if (!pivotTable || fields.length === 0) return;

        document.removeEventListener('mousedown', this._boundClickOutside);
        this._createPanel();

        const pendingSelections = new Map();
        fields.forEach(fieldIndex => {
            const field = pivotTable.pivotFields?.[fieldIndex];
            const selected = new Set();
            const filterableItems = this._getPivotFilterableItems(field);
            filterableItems.forEach(({ item, itemIndex }) => {
                if (item.h !== '1') selected.add(String(itemIndex));
            });
            pendingSelections.set(fieldIndex, selected);
        });

        this._customContext = {
            type: 'pivot',
            pivotTable,
            fields,
            activeField: fields[0],
            pendingSelections
        };
        this._showCount = false;
        this._filterMode = 'list';
        this._dateTree = null;
        this._collapsedDateNodes = new Set();
        this._searchKeyword = '';
        this._toggleFieldSelect(fields.length > 1, fields);
        this._setPivotField(fields[0]);
        this._positionPanel(x, y);
    }

    hide() {
        if (this._panel) {
            this._panel.classList.remove('active');
        }
        this._currentColIndex = null;
        this._currentScopeId = null;
        this._customContext = null;
        this._filterMode = 'list';
        this._dateTree = null;
        this._collapsedDateNodes = new Set();
        this._searchKeyword = '';
        document.removeEventListener('mousedown', this._boundClickOutside);
    }

    destroy() {
        this.hide();
        if (this._panel && this._panel.parentNode) {
            this._panel.parentNode.removeChild(this._panel);
        }
        this._panel = null;
    }

    _renderList(values) {
        const list = this._panel.querySelector('.sn-filter-list');
        if (!list) return;

        const allowMultiple = this._customContext?.type !== 'pivot'
            || this._customContext?.pivotTable?.multipleFieldFilters !== false;
        if (!allowMultiple && this._selectedValues.size > 1) {
            const first = this._selectedValues.values().next().value;
            this._selectedValues = new Set(first ? [first] : []);
        }

        list.innerHTML = values.map(item => {
            const key = item.filterKey ?? String(item.value ?? '');
            const checked = this._selectedValues.has(key) ? 'checked' : '';
            const text = this._escapeHtml(item.text);
            const showCount = this._showCount && item.count !== undefined && item.count !== null;
            return `
                <div class="sn-filter-item" data-value="${this._escapeHtml(key)}">
                    <input type="checkbox" ${checked}>
                    <label>${text}</label>
                    ${showCount ? `<span class="sn-filter-count">(${item.count})</span>` : ''}
                </div>
            `;
        }).join('');

        list.querySelectorAll('.sn-filter-item').forEach(item => {
            const checkbox = item.querySelector('input[type="checkbox"]');
            const value = item.dataset.value;

            item.addEventListener('click', (e) => {
                if (e.target !== checkbox) {
                    checkbox.checked = !checkbox.checked;
                }
                if (!allowMultiple) {
                    this._selectedValues.clear();
                    this._selectedValues.add(value);
                    this._renderCurrentView();
                    return;
                }
                if (checkbox.checked) {
                    this._selectedValues.add(value);
                } else {
                    this._selectedValues.delete(value);
                }
            });
        });
    }

    _renderCurrentView() {
        if (this._customContext?.type === 'pivot') {
            this._renderList(this._allValues);
            return;
        }

        if (this._filterMode === 'date' && !this._searchKeyword) {
            this._renderDateTree();
            return;
        }

        const values = this._searchKeyword
            ? this._filterValuesByKeyword(this._searchKeyword)
            : this._allValues;
        this._renderList(values);
    }

    _filterValuesByKeyword(keyword) {
        const normalized = String(keyword ?? '').toLowerCase().trim();
        if (!normalized) return this._allValues;
        return this._allValues.filter(item => {
            const displayText = String(item.text ?? '').toLowerCase();
            const dateLabel = String(item.dateLabel ?? '').toLowerCase();
            return displayText.includes(normalized) || dateLabel.includes(normalized);
        });
    }

    _buildDateTree(values) {
        if (!Array.isArray(values) || values.length === 0) return null;

        const yearMap = new Map();
        const otherValues = [];
        let dateCount = 0;
        let otherCount = 0;

        const ensureNode = (map, key, createFn) => {
            if (!map.has(key)) map.set(key, createFn());
            return map.get(key);
        };

        values.forEach(item => {
            if (!item || item.isEmpty) {
                otherValues.push(item);
                return;
            }

            if (!item.isDateValue || !item.dateKey) {
                otherValues.push(item);
                otherCount += item.count ?? 0;
                return;
            }

            dateCount += item.count ?? 0;
            const filterKey = item.filterKey ?? String(item.value ?? '');
            const yearKey = String(item.dateYear);
            const monthKey = `${yearKey}-${String(item.dateMonth).padStart(2, '0')}`;

            const yearNode = ensureNode(yearMap, yearKey, () => ({
                id: `year:${yearKey}`,
                label: yearKey,
                level: 0,
                count: 0,
                keys: [],
                months: new Map()
            }));
            yearNode.count += item.count ?? 0;
            yearNode.keys.push(filterKey);

            const monthNode = ensureNode(yearNode.months, monthKey, () => ({
                id: `month:${monthKey}`,
                label: monthKey,
                level: 1,
                count: 0,
                keys: [],
                days: new Map()
            }));
            monthNode.count += item.count ?? 0;
            monthNode.keys.push(filterKey);

            const dayNode = ensureNode(monthNode.days, item.dateKey, () => ({
                id: `day:${item.dateKey}`,
                label: item.dateLabel || item.dateKey,
                level: 2,
                count: 0,
                keys: []
            }));
            dayNode.count += item.count ?? 0;
            dayNode.keys.push(filterKey);
        });

        if (dateCount === 0 || dateCount <= otherCount) {
            return null;
        }

        const sortAsc = (a, b) => String(a).localeCompare(String(b));
        const roots = Array.from(yearMap.keys()).sort(sortAsc).map(yearKey => {
            const yearNode = yearMap.get(yearKey);
            yearNode.children = Array.from(yearNode.months.keys()).sort(sortAsc).map(monthKey => {
                const monthNode = yearNode.months.get(monthKey);
                monthNode.children = Array.from(monthNode.days.keys()).sort(sortAsc).map(dayKey => monthNode.days.get(dayKey));
                delete monthNode.days;
                return monthNode;
            });
            delete yearNode.months;
            return yearNode;
        });

        const nodeMap = new Map();
        const walk = node => {
            nodeMap.set(node.id, node);
            node.children?.forEach(walk);
        };
        roots.forEach(walk);

        return { roots, nodeMap, otherValues };
    }

    _renderDateTree() {
        const list = this._panel.querySelector('.sn-filter-list');
        if (!list || !this._dateTree) return;

        const { roots, otherValues } = this._dateTree;
        const parts = [];

        if (roots.length > 0) {
            parts.push(`<div class="sn-filter-tree-section-title">${this._t('autoFilter.panel.dateSection')}</div>`);
            parts.push(roots.map(node => this._renderDateNode(node)).join(''));
        }

        if (otherValues.length > 0) {
            parts.push(`<div class="sn-filter-tree-section-title">${this._t('autoFilter.panel.otherValues')}</div>`);
            parts.push(otherValues.map(item => {
                const key = item.filterKey ?? String(item.value ?? '');
                const checked = this._selectedValues.has(key) ? 'checked' : '';
                const text = this._escapeHtml(item.text);
                const showCount = this._showCount && item.count !== undefined && item.count !== null;
                return `
                    <div class="sn-filter-item sn-filter-mixed-item" data-value="${this._escapeHtml(key)}">
                        <input type="checkbox" ${checked}>
                        <label>${text}</label>
                        ${showCount ? `<span class="sn-filter-count">(${item.count})</span>` : ''}
                    </div>
                `;
            }).join(''));
        }

        list.innerHTML = parts.join('');

        list.querySelectorAll('.sn-filter-tree-item').forEach(itemEl => {
            const node = this._dateTree.nodeMap.get(itemEl.dataset.nodeId);
            const checkbox = itemEl.querySelector('input[type="checkbox"]');
            const selectionState = this._getSelectionState(node?.keys || []);
            checkbox.checked = selectionState.checked;
            checkbox.indeterminate = selectionState.indeterminate;

            itemEl.addEventListener('click', (e) => {
                if (e.target.closest('.sn-filter-tree-toggle')) {
                    return;
                }
                if (e.target !== checkbox) {
                    checkbox.checked = !checkbox.checked;
                    checkbox.indeterminate = false;
                }
                this._setKeysSelected(node?.keys || [], checkbox.checked);
                this._renderCurrentView();
            });
        });

        list.querySelectorAll('.sn-filter-tree-toggle').forEach(toggleEl => {
            toggleEl.addEventListener('click', (e) => {
                e.stopPropagation();
                this._toggleDateNode(toggleEl.dataset.nodeId);
            });
        });

        list.querySelectorAll('.sn-filter-mixed-item').forEach(itemEl => {
            const checkbox = itemEl.querySelector('input[type="checkbox"]');
            const value = itemEl.dataset.value;

            itemEl.addEventListener('click', (e) => {
                if (e.target !== checkbox) {
                    checkbox.checked = !checkbox.checked;
                }
                if (checkbox.checked) {
                    this._selectedValues.add(value);
                } else {
                    this._selectedValues.delete(value);
                }
            });
        });
    }

    _renderDateNode(node) {
        const hasChildren = Array.isArray(node.children) && node.children.length > 0;
        const collapsed = hasChildren && this._collapsedDateNodes.has(node.id);
        return `
            <div class="sn-filter-tree-group">
                <div class="sn-filter-item sn-filter-tree-item level-${node.level}" data-node-id="${node.id}">
                    ${hasChildren
                        ? `<button type="button" class="sn-filter-tree-toggle ${collapsed ? 'collapsed' : ''}" data-node-id="${node.id}" aria-label="toggle"></button>`
                        : '<span class="sn-filter-tree-toggle-placeholder"></span>'}
                    <input type="checkbox">
                    <label>${this._escapeHtml(node.label)}</label>
                    <span class="sn-filter-count">(${node.count})</span>
                </div>
                ${hasChildren && !collapsed ? `<div class="sn-filter-tree-children">${node.children.map(child => this._renderDateNode(child)).join('')}</div>` : ''}
            </div>
        `;
    }

    _initDateTreeCollapsedState() {
        this._collapsedDateNodes = new Set();
        if (!this._dateTree?.roots?.length) return;

        const collapseRoots = this._dateTree.roots.length > 1;
        this._dateTree.roots.forEach(root => {
            if (collapseRoots) {
                this._collapsedDateNodes.add(root.id);
            }
            root.children?.forEach(child => {
                this._collapsedDateNodes.add(child.id);
            });
        });
    }

    _toggleDateNode(nodeId) {
        if (!nodeId) return;
        if (this._collapsedDateNodes.has(nodeId)) {
            this._collapsedDateNodes.delete(nodeId);
        } else {
            this._collapsedDateNodes.add(nodeId);
        }
        this._renderCurrentView();
    }

    _getSelectionState(keys) {
        if (!Array.isArray(keys) || keys.length === 0) {
            return { checked: false, indeterminate: false };
        }
        let selectedCount = 0;
        keys.forEach(key => {
            if (this._selectedValues.has(key)) selectedCount++;
        });
        return {
            checked: selectedCount === keys.length,
            indeterminate: selectedCount > 0 && selectedCount < keys.length
        };
    }

    _setKeysSelected(keys, checked) {
        if (!Array.isArray(keys)) return;
        keys.forEach(key => {
            if (checked) {
                this._selectedValues.add(key);
            } else {
                this._selectedValues.delete(key);
            }
        });
    }

    _onSearch(keyword) {
        this._searchKeyword = String(keyword ?? '').toLowerCase().trim();
        this._renderCurrentView();
    }

    _selectAll() {
        if (this._customContext?.type === 'pivot') {
            const fieldIndex = this._customContext.activeField;
            const pendingSelections = this._customContext.pendingSelections;
            let selected = pendingSelections?.get(fieldIndex);
            if (!selected) {
                selected = new Set();
                pendingSelections?.set(fieldIndex, selected);
            }
            selected.clear();
            this._allValues.forEach(item => {
                selected.add(String(item.value ?? ''));
            });
            this._selectedValues = selected;
            this._renderCurrentView();
            return;
        }

        this._selectedValues = new Set(this._allValues.map(v => v.filterKey ?? String(v.value ?? '')));
        this._renderCurrentView();
    }

    _clearAll() {
        if (this._customContext?.type === 'pivot') {
            const fieldIndex = this._customContext.activeField;
            const pendingSelections = this._customContext.pendingSelections;
            let selected = pendingSelections?.get(fieldIndex);
            if (!selected) {
                selected = new Set();
                pendingSelections?.set(fieldIndex, selected);
            }
            selected.clear();
            this._selectedValues = selected;
            this._renderCurrentView();
            return;
        }

        this._selectedValues.clear();
        this._renderCurrentView();
    }

    _doSort(type) {
        if (this._customContext?.type === 'pivot') {
            const pivotTable = this._customContext.pivotTable;
            const fieldIndex = this._customContext.activeField;
            if (!pivotTable || fieldIndex === null || fieldIndex === undefined) return;

            pivotTable.setFieldSort(fieldIndex, type === 'asc' ? 'ascending' : 'descending');
            pivotTable.calculate();
            pivotTable.render();
            this.hide();
            return;
        }

        if (this._currentColIndex === null) return;

        const sheet = this._SN.activeSheet;
        const range = this.autoFilter.getScopeRange(this._currentScopeId);
        if (!range) return;

        const table = typeof this._currentScopeId === 'string' && this._currentScopeId.startsWith('table:')
            ? sheet.Table.get(this._currentScopeId.slice(6))
            : null;
        const success = sheet.rangeSort(
            [{ col: this._currentColIndex, order: type }],
            range,
            { hasHeader: table ? table.showHeaderRow === true : true }
        );

        if (success) {
            this.autoFilter.setSortState(this._currentColIndex, type, this._currentScopeId);
            this.autoFilter._applyFilters(this._currentScopeId);
        }

        this.hide();
        this._SN._r();
    }

    _apply() {
        if (this._customContext?.type === 'pivot') {
            const pivotTable = this._customContext.pivotTable;
            const fields = this._customContext.fields || [];
            const pendingSelections = this._customContext.pendingSelections;
            if (!pivotTable || fields.length === 0 || !pendingSelections) return;

            fields.forEach(fieldIndex => {
                const field = pivotTable.pivotFields?.[fieldIndex];
                if (!field) return;
                const selected = pendingSelections.get(fieldIndex) || new Set();
                const filterableItems = this._getPivotFilterableItems(field);
                if (filterableItems.length === 0) {
                    pivotTable.clearFieldFilter(fieldIndex);
                    return;
                }

                if (selected.size === filterableItems.length) {
                    pivotTable.clearFieldFilter(fieldIndex);
                    return;
                }

                const hiddenMap = {};
                filterableItems.forEach(({ itemIndex }) => {
                    hiddenMap[itemIndex] = !selected.has(String(itemIndex));
                });
                pivotTable.setItemsHidden(fieldIndex, hiddenMap);
            });

            pivotTable.calculate();
            pivotTable.render();
            this.hide();
            return;
        }

        if (this._currentColIndex === null || !this.autoFilter) return;

        if (this._selectedValues.size === this._allValues.length) {
            this.autoFilter.clearColumnFilter(this._currentColIndex, this._currentScopeId);
        } else {
            const values = Array.from(this._selectedValues).map(strVal => {
                const item = this._allValues.find(v => (v.filterKey ?? String(v.value ?? '')) === strVal);
                return item ? item.value : strVal;
            });
            this.autoFilter.setColumnFilter(this._currentColIndex, {
                type: 'values',
                values
            }, this._currentScopeId);
        }

        this.hide();
    }

    _toggleFieldSelect(show, fields = []) {
        const fieldWrap = this._panel?.querySelector('.sn-filter-field');
        const fieldSelect = this._panel?.querySelector('.sn-filter-field-select');
        if (!fieldWrap || !fieldSelect) return;

        if (!show) {
            fieldWrap.classList.remove('active');
            fieldSelect.innerHTML = '';
            return;
        }

        const pivotTable = this._customContext?.pivotTable;
        fieldSelect.innerHTML = fields.map(fieldIndex => {
            const name = pivotTable?.pivotFields?.[fieldIndex]?.name || this._t('pivot.field.fallbackName', { index: fieldIndex + 1 });
            return `<option value="${fieldIndex}">${name}</option>`;
        }).join('');
        fieldWrap.classList.add('active');
        fieldSelect.value = String(this._customContext?.activeField ?? fields[0]);
    }

    _setPivotField(fieldIndex) {
        const pivotTable = this._customContext?.pivotTable;
        if (!pivotTable) return;
        const field = pivotTable.pivotFields?.[fieldIndex];
        if (!field) return;

        const filterableItems = this._getPivotFilterableItems(field);
        const availableIndexes = new Set(filterableItems.map(({ itemIndex }) => String(itemIndex)));

        this._customContext.activeField = fieldIndex;
        this._allValues = filterableItems.map(({ item, itemIndex }) => ({
            value: String(itemIndex),
            filterKey: String(itemIndex),
            text: pivotTable._getItemDisplayValue(fieldIndex, item.x) || ''
        }));

        const pendingSelections = this._customContext?.pendingSelections;
        let selected = pendingSelections?.get(fieldIndex);
        if (!selected) {
            selected = new Set();
            filterableItems.forEach(({ item, itemIndex }) => {
                if (item.h !== '1') selected.add(String(itemIndex));
            });
            pendingSelections?.set(fieldIndex, selected);
        } else {
            selected = new Set([...selected].filter(value => availableIndexes.has(String(value))));
            pendingSelections?.set(fieldIndex, selected);
        }
        this._selectedValues = selected;

        const title = this._panel?.querySelector('.sn-filter-title');
        if (title) {
            const fieldLabel = field.name || this._t('pivot.field.fallbackName', { index: fieldIndex + 1 });
            title.textContent = this._t('autoFilter.panel.titleField', { field: fieldLabel });
        }

        const searchInput = this._panel?.querySelector('.sn-filter-search-input');
        if (searchInput) searchInput.value = '';

        this._panel?.querySelectorAll('.sn-filter-sort-item').forEach(item => {
            const sortType = field.sortType;
            const isActive = (item.dataset.sort === 'asc' && sortType === 'ascending')
                || (item.dataset.sort === 'desc' && sortType === 'descending');
            item.classList.toggle('active', isActive);
        });

        this._renderCurrentView();
    }

    _getPivotFilterableItems(field) {
        if (!field?.items?.length) return [];
        const result = [];
        field.items.forEach((item, itemIndex) => {
            if (!Number.isFinite(item?.x)) return;
            result.push({ item, itemIndex });
        });
        return result;
    }

    _positionPanel(x, y) {
        const panelWidth = 260;
        const panelHeight = 400;
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;

        const canvasRect = this.canvas.handleLayer.getBoundingClientRect();
        let left = canvasRect.left + x;
        let top = canvasRect.top + y;

        if (left + panelWidth > viewportWidth - 10) {
            left = viewportWidth - panelWidth - 10;
        }
        if (top + panelHeight > viewportHeight - 10) {
            top = viewportHeight - panelHeight - 10;
        }
        if (left < 10) left = 10;
        if (top < 10) top = 10;

        this._panel.style.position = 'fixed';
        this._panel.style.left = `${left}px`;
        this._panel.style.top = `${top}px`;
        this._panel.style.zIndex = '10000';
        this._panel.classList.add('active');

        setTimeout(() => {
            document.addEventListener('mousedown', this._boundClickOutside);
        }, 0);
    }

    _escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }
}
