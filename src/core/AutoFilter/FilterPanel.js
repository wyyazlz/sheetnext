/**
 * 筛选面板交互模块
 * 负责筛选面板的显示、隐藏和用户交互
 * 采用按需创建DOM的方式，避免overflow:hidden等问题
 */

import { sortRange } from '../Utils/SortUtils.js';
import getSvg from '../../assets/mainSvgs.js';

export default class FilterPanel {
    constructor(canvas) {
        this.canvas = canvas;
        this._SN = canvas.SN;
        this._panel = null; // 按需创建
        this._currentColIndex = null; // 当前正在筛选的列
        this._currentScopeId = null; // 当前筛选作用域
        this._allValues = []; // 所有可选值
        this._selectedValues = new Set(); // 选中的值
        this._boundClickOutside = this._onClickOutside.bind(this);
        this._customContext = null;
        this._showCount = true;
    }

    /**
     * 获取当前活动工作表的筛选器
     */
    get autoFilter() {
        return this._SN.activeSheet?.AutoFilter;
    }

    _t(key, params = {}) {
        if (!key || typeof key !== 'string') return '';
        return this._SN.t(key, params);
    }

    /**
     * 创建筛选面板DOM
     */
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

        // 挂载到body确保不受父容器overflow影响
        document.body.appendChild(panel);
        this._panel = panel;
        this._bindEvents();

        return panel;
    }

    /**
     * 绑定面板事件
     */
    _bindEvents() {
        const panel = this._panel;
        if (!panel) return;

        // 关闭按钮
        panel.querySelector('.sn-filter-close')?.addEventListener('click', () => this.hide());

        // 搜索框
        const searchInput = panel.querySelector('.sn-filter-search-input');
        searchInput?.addEventListener('input', (e) => this._onSearch(e.target.value));

        // 全选/清除
        panel.querySelector('.sn-filter-select-all')?.addEventListener('click', () => this._selectAll());
        panel.querySelector('.sn-filter-clear-all')?.addEventListener('click', () => this._clearAll());

        // 排序
        panel.querySelectorAll('.sn-filter-sort-item').forEach(item => {
            item.addEventListener('click', () => {
                const sortType = item.dataset.sort;
                this._doSort(sortType);
            });
        });

        // 确定/取消
        panel.querySelector('.sn-filter-apply')?.addEventListener('click', () => this._apply());
        panel.querySelector('.sn-filter-cancel')?.addEventListener('click', () => this.hide());

        // 阻止面板内的点击冒泡
        panel.addEventListener('mousedown', (e) => e.stopPropagation());

        const fieldSelect = panel.querySelector('.sn-filter-field-select');
        fieldSelect?.addEventListener('change', () => {
            if (this._customContext?.type === 'pivot') {
                const fieldIndex = Number(fieldSelect.value);
                this._setPivotField(fieldIndex);
            }
        });
    }

    /**
     * 点击外部关闭处理
     */
    _onClickOutside(e) {
        if (!this._panel) return;
        // 检查是否点击了筛选图标
        const isFilterIcon = e.target.closest && e.target.closest('.sn-filter-icon');
        if (!isFilterIcon && !this._panel.contains(e.target)) {
            this.hide();
        }
    }

    /**
     * 显示筛选面板
     * @param {number} colIndex - 列索引
     * @param {number} x - 面板X位置（相对于视口）
     * @param {number} y - 面板Y位置（相对于视口）
     */
    show(colIndex, x, y, scopeId = null) {
        const autoFilter = this.autoFilter;
        if (!autoFilter?.isScopeEnabled(scopeId)) return;

        // 先移除旧的点击外部监听器，避免事件冒泡时立即关闭新面板
        document.removeEventListener('mousedown', this._boundClickOutside);

        // 确保面板已创建
        this._createPanel();

        this._currentColIndex = colIndex;
        this._currentScopeId = scopeId;
        this._customContext = null;
        this._showCount = true;
        this._toggleFieldSelect(false);

        // 获取列的所有值
        this._allValues = autoFilter.getColumnValues(colIndex, this._currentScopeId);

        // 获取当前筛选条件
        const currentFilter = autoFilter.getColumnFilter(colIndex, this._currentScopeId);
        if (currentFilter?.type === 'values') {
            this._selectedValues = new Set(currentFilter.values.map(v => String(v ?? '')));
        } else {
            // 默认全选
            this._selectedValues = new Set(this._allValues.map(v => String(v.value ?? '')));
        }

        // 渲染列表
        this._renderList(this._allValues);

        // 清空搜索框
        const searchInput = this._panel.querySelector('.sn-filter-search-input');
        if (searchInput) searchInput.value = '';

        // 更新标题
        const title = this._panel.querySelector('.sn-filter-title');
        if (title) {
            const colName = this._SN.Utils.numToChar(colIndex);
            title.textContent = this._t('autoFilter.panel.titleColumn', { column: colName });
        }

        // 更新排序按钮选中状态
        const currentSort = autoFilter.getSortOrder(colIndex, this._currentScopeId);
        this._panel.querySelectorAll('.sn-filter-sort-item').forEach(item => {
            if (item.dataset.sort === currentSort) {
                item.classList.add('active');
            } else {
                item.classList.remove('active');
            }
        });

        this._positionPanel(x, y);
    }

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
        this._toggleFieldSelect(fields.length > 1, fields);
        this._setPivotField(fields[0]);
        this._positionPanel(x, y);
    }

    /**
     * 隐藏筛选面板
     */
    hide() {
        if (this._panel) {
            this._panel.classList.remove('active');
        }
        this._currentColIndex = null;
        this._currentScopeId = null;
        this._customContext = null;
        document.removeEventListener('mousedown', this._boundClickOutside);
    }

    /**
     * 销毁面板（释放资源）
     */
    destroy() {
        this.hide();
        if (this._panel && this._panel.parentNode) {
            this._panel.parentNode.removeChild(this._panel);
        }
        this._panel = null;
    }

    /**
     * 渲染值列表
     */
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
            const key = String(item.value ?? '');
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

        // 绑定checkbox事件
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
                    this._renderList(this._allValues);
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

    /**
     * 搜索过滤
     */
    _onSearch(keyword) {
        keyword = keyword.toLowerCase().trim();
        const filtered = keyword
            ? this._allValues.filter(v => v.text.toLowerCase().includes(keyword))
            : this._allValues;
        this._renderList(filtered);
    }

    /**
     * 全选
     */
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
            this._renderList(this._allValues);
            return;
        }

        this._selectedValues = new Set(this._allValues.map(v => String(v.value ?? '')));
        this._renderList(this._allValues);
    }

    /**
     * 清除选择
     */
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
            this._renderList(this._allValues);
            return;
        }

        this._selectedValues.clear();
        this._renderList(this._allValues);
    }

    /**
     * 执行排序
     */
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

        // 使用通用排序（hasHeader 跳过表头）
        const success = sortRange(sheet, range, { col: this._currentColIndex, order: type }, { hasHeader: true });

        if (success) {
            this.autoFilter.setSortState(this._currentColIndex, type, this._currentScopeId);
            this.autoFilter._applyFilters(this._currentScopeId);
        }

        this.hide();
        this._SN._r();
    }

    /**
     * 应用筛选
     */
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

        // 如果全选，则清除筛选条件
        if (this._selectedValues.size === this._allValues.length) {
            this.autoFilter.clearColumnFilter(this._currentColIndex, this._currentScopeId);
        } else {
            // 转换选中值为正确类型
            const values = Array.from(this._selectedValues).map(strVal => {
                const item = this._allValues.find(v => String(v.value ?? '') === strVal);
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
            const isActive = (item.dataset.sort === 'asc' && sortType === 'ascending') ||
                (item.dataset.sort === 'desc' && sortType === 'descending');
            item.classList.toggle('active', isActive);
        });

        this._renderList(this._allValues);
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

    /**
     * HTML转义
     */
    _escapeHtml(str) {
        if (str === null || str === undefined) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }
}
