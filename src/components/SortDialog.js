/**
 * 自定义排序对话框
 * 复刻 Excel 完整排序能力：多级排序、排序依据、自定义列表等
 */

import { sortRange } from '../core/Utils/SortUtils.js';
import getSvg from '../assets/mainSvgs.js';

// 预设的自定义列表
const PRESET_CUSTOM_LISTS = [
    { name: '星期', values: ['周一', '周二', '周三', '周四', '周五', '周六', '周日'] },
    { name: '月份', values: ['一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'] },
    { name: '优先级', values: ['高', '中', '低'] },
    { name: '季度', values: ['第一季度', '第二季度', '第三季度', '第四季度'] }
];

export default class SortDialog {
    constructor(SN) {
        this.SN = SN;
        this.overlay = null;
        this.sortLevels = [];      // 排序条件列表
        this.options = {
            hasHeader: false,      // 默认数据不包含标题
            caseSensitive: false   // 区分大小写
        };
        this.range = null;         // 排序范围
        this.columns = [];         // 可选列列表
        this.customLists = [...PRESET_CUSTOM_LISTS]; // 自定义列表（包含预设）
    }

    _t(key, params) {
        return this.SN.t(key, params);
    }

    /**
     * 显示排序对话框
     */
    show() {
        const sheet = this.SN.activeSheet;

        // 获取排序范围
        this.range = this._getRange(sheet);
        if (!this.range) {
            this.SN.Utils.toast(this.SN.t('components.sortDialog.toast.msg001'));
            return;
        }

        // 检查合并单元格
        if (sheet.areaHaveMerge(this.range)) {
            this.SN.Utils.toast(this.SN.t('components.sortDialog.toast.msg002'));
            return;
        }

        // 初始化列列表
        this._initColumns();

        // 初始化默认排序条件
        if (this.sortLevels.length === 0) {
            this.sortLevels = [{
                column: this.range.s.c,
                sortBy: 'value',
                order: 'asc',
                customOrder: null
            }];
        }

        this._createDialog();
        this._show();
    }

    /**
     * 获取排序范围
     */
    _getRange(sheet) {
        const areas = sheet.activeAreas;
        if (areas.length > 0) {
            const area = areas[areas.length - 1];
            if (area.e.r > area.s.r || area.e.c > area.s.c) {
                return area;
            }
        }
        return this._detectDataRange(sheet);
    }

    /**
     * 自动检测数据区域
     */
    _detectDataRange(sheet) {
        const startCell = sheet.activeCell;
        let minR = startCell.r, maxR = startCell.r;
        let minC = startCell.c, maxC = startCell.c;

        const hasData = (r, c) => {
            const cell = sheet.getCell(r, c);
            return cell.editVal !== undefined && cell.editVal !== null && cell.editVal !== '';
        };

        while (minR > 0 && hasData(minR - 1, startCell.c)) minR--;
        while (maxR < sheet.rowCount - 1 && hasData(maxR + 1, startCell.c)) maxR++;
        while (minC > 0 && hasData(startCell.r, minC - 1)) minC--;
        while (maxC < sheet.colCount - 1 && hasData(startCell.r, maxC + 1)) maxC++;

        if (minR === maxR && minC === maxC) return null;

        return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC } };
    }

    /**
     * 初始化列列表
     */
    _initColumns() {
        const sheet = this.SN.activeSheet;
        this.columns = [];

        for (let c = this.range.s.c; c <= this.range.e.c; c++) {
            const headerCell = sheet.getCell(this.range.s.r, c);
            const headerVal = headerCell.showVal ?? headerCell.editVal ?? '';
            const colLetter = this.SN.Utils.numToChar(c);

            this.columns.push({
                index: c,
                letter: colLetter,
                name: this.options.hasHeader && headerVal
                    ? String(headerVal)
                    : this._t('components.sortDialog.content.columnFallback', { colLetter })
            });
        }
    }

    /**
     * 创建对话框
     */
    _createDialog() {
        if (this.overlay) {
            this._updateContent();
            return;
        }

        const overlay = document.createElement('div');
        overlay.className = 'sn-modal-overlay sn-sort-dialog-overlay';
        overlay.innerHTML = `
            <div class="sn-modal sn-sort-dialog">
                <div class="sn-modal-header">
                    <h3 class="sn-modal-title">${this._t('components.sortDialog.content.title')}</h3>
                    <button class="sn-modal-close" title="${this._t('components.sortDialog.content.close')}"></button>
                </div>
                <div class="sn-modal-body">
                    <div class="sn-sort-toolbar">
                        <button class="sn-sort-btn" data-action="add">
                            ${getSvg('sort_add')}
                            ${this._t('components.sortDialog.content.toolbar.add')}
                        </button>
                        <button class="sn-sort-btn" data-action="delete" disabled>
                            ${getSvg('sort_remove')}
                            ${this._t('components.sortDialog.content.toolbar.delete')}
                        </button>
                        <button class="sn-sort-btn" data-action="copy" disabled>
                            ${getSvg('sort_copy')}
                            ${this._t('components.sortDialog.content.toolbar.copy')}
                        </button>
                        <div class="sn-sort-toolbar-divider"></div>
                        <button class="sn-sort-btn" data-action="moveUp" disabled>
                            ${getSvg('sort_moveUp')}
                        </button>
                        <button class="sn-sort-btn" data-action="moveDown" disabled>
                            ${getSvg('sort_moveDown')}
                        </button>
                    </div>
                    <div class="sn-sort-levels-header">
                        <span class="sn-sort-col-select"></span>
                        <span class="sn-sort-col-column">${this._t('components.sortDialog.content.header.column')}</span>
                        <span class="sn-sort-col-sortby">${this._t('components.sortDialog.content.header.sortBy')}</span>
                        <span class="sn-sort-col-order">${this._t('components.sortDialog.content.header.order')}</span>
                    </div>
                    <div class="sn-sort-levels"></div>
                    <div class="sn-sort-options">
                        <label class="sn-sort-option">
                            <input type="checkbox" class="sn-sort-has-header" ${this.options.hasHeader ? 'checked' : ''}>
                            <span>${this._t('components.sortDialog.content.options.hasHeader')}</span>
                        </label>
                        <label class="sn-sort-option">
                            <input type="checkbox" class="sn-sort-case-sensitive" ${this.options.caseSensitive ? 'checked' : ''}>
                            <span>${this._t('components.sortDialog.content.options.caseSensitive')}</span>
                        </label>
                    </div>
                </div>
                <div class="sn-modal-footer">
                    <button class="sn-modal-btn sn-modal-btn-cancel">${this._t('components.sortDialog.content.cancel')}</button>
                    <button class="sn-modal-btn sn-modal-btn-confirm">${this._t('components.sortDialog.content.confirm')}</button>
                </div>
            </div>
        `;

        document.body.appendChild(overlay);
        this.overlay = overlay;

        this._bindEvents();
        this._updateContent();
    }

    /**
     * 绑定事件
     */
    _bindEvents() {
        const overlay = this.overlay;

        overlay.querySelector('.sn-modal-close').addEventListener('click', () => this.hide());
        overlay.querySelector('.sn-modal-btn-cancel').addEventListener('click', () => this.hide());
        overlay.querySelector('.sn-modal-btn-confirm').addEventListener('click', () => this._apply());

        overlay.querySelectorAll('.sn-sort-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this._handleToolbarAction(btn.dataset.action);
            });
        });

        overlay.querySelector('.sn-sort-has-header').addEventListener('change', (e) => {
            this.options.hasHeader = e.target.checked;
            this._initColumns();
            this._updateContent();
        });

        overlay.querySelector('.sn-sort-case-sensitive').addEventListener('change', (e) => {
            this.options.caseSensitive = e.target.checked;
        });

        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) this.hide();
        });

        this._escHandler = (e) => {
            if (e.key === 'Escape') {
                // 如果自定义列表弹窗打开，先关闭它
                const customListPopup = overlay.querySelector('.sn-sort-custom-list-popup.active');
                if (customListPopup) {
                    customListPopup.classList.remove('active');
                } else {
                    this.hide();
                }
            }
        };
        document.addEventListener('keydown', this._escHandler);
    }

    /**
     * 更新内容
     */
    _updateContent() {
        const levelsContainer = this.overlay.querySelector('.sn-sort-levels');
        levelsContainer.innerHTML = this.sortLevels.map((level, index) => this._renderLevel(level, index)).join('');

        levelsContainer.querySelectorAll('.sn-sort-level').forEach((row, index) => {
            row.addEventListener('click', (e) => {
                if (e.target.tagName !== 'SELECT' && e.target.tagName !== 'BUTTON' && !e.target.closest('.sn-sort-custom-list-popup')) {
                    this._selectLevel(index);
                }
            });

            row.querySelector('.sn-sort-column-select').addEventListener('change', (e) => {
                this.sortLevels[index].column = parseInt(e.target.value);
            });

            row.querySelector('.sn-sort-sortby-select').addEventListener('change', (e) => {
                this.sortLevels[index].sortBy = e.target.value;
                this._updateOrderOptions(row, index);
            });

            const orderSelect = row.querySelector('.sn-sort-order-select');
            orderSelect.addEventListener('change', (e) => {
                const value = e.target.value;
                if (value === 'custom') {
                    this._showCustomListPopup(row, index);
                } else {
                    this.sortLevels[index].order = value;
                    this.sortLevels[index].customOrder = null;
                }
            });

            // 编辑自定义列表按钮
            const editBtn = row.querySelector('.sn-sort-edit-custom');
            if (editBtn) {
                editBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this._showCustomListPopup(row, index);
                });
            }
        });

        this._updateToolbarState();
    }

    /**
     * 渲染单个排序条件
     */
    _renderLevel(level, index) {
        const isSelected = index === this.selectedIndex;
        const prefix = index === 0
            ? this._t('components.sortDialog.content.level.primary')
            : (index === 1
                ? this._t('components.sortDialog.content.level.secondary')
                : this._t('components.sortDialog.content.level.nth', { index: index + 1 }));

        const columnOptions = this.columns.map(col =>
            `<option value="${col.index}" ${col.index === level.column ? 'selected' : ''}>${col.name}</option>`
        ).join('');

        const sortByOptions = `
            <option value="value" ${level.sortBy === 'value' ? 'selected' : ''}>${this._t('components.sortDialog.content.sortBy.value')}</option>
            <option value="cellColor" ${level.sortBy === 'cellColor' ? 'selected' : ''}>${this._t('components.sortDialog.content.sortBy.cellColor')}</option>
            <option value="fontColor" ${level.sortBy === 'fontColor' ? 'selected' : ''}>${this._t('components.sortDialog.content.sortBy.fontColor')}</option>
        `;

        const orderOptions = this._getOrderOptions(level);

        // 如果是自定义列表，显示编辑按钮
        const hasCustomOrder = level.customOrder && level.customOrder.length > 0;
        const customEditBtn = hasCustomOrder ?
            `<button class="sn-sort-edit-custom" title="${this._t('components.sortDialog.content.customList.editTitle')}">...</button>` : '';

        return `
            <div class="sn-sort-level ${isSelected ? 'selected' : ''}" data-index="${index}">
                <span class="sn-sort-col-select">
                    <input type="radio" name="sn-sort-select" ${isSelected ? 'checked' : ''}>
                </span>
                <span class="sn-sort-col-column">
                    <span class="sn-sort-level-prefix">${prefix}</span>
                    <select class="sn-sort-column-select">${columnOptions}</select>
                </span>
                <span class="sn-sort-col-sortby">
                    <select class="sn-sort-sortby-select">${sortByOptions}</select>
                </span>
                <span class="sn-sort-col-order">
                    <select class="sn-sort-order-select">${orderOptions}</select>
                    ${customEditBtn}
                </span>
            </div>
        `;
    }

    /**
     * 获取次序选项
     */
    _getOrderOptions(level) {
        const { sortBy, order, customOrder } = level;

        if (sortBy === 'value') {
            const hasCustom = Array.isArray(customOrder) && customOrder.length > 0;
            const customPreview = hasCustom ? customOrder.slice(0, 3).join(', ') : '';
            const customLabel = hasCustom
                ? this._t('components.sortDialog.content.order.customPrefix', {
                    preview: `${customPreview}${customOrder.length > 3 ? '...' : ''}`
                })
                : this._t('components.sortDialog.content.order.customList');

            return `
                <option value="asc" ${order === 'asc' && !hasCustom ? 'selected' : ''}>${this._t('components.sortDialog.content.order.asc')}</option>
                <option value="desc" ${order === 'desc' && !hasCustom ? 'selected' : ''}>${this._t('components.sortDialog.content.order.desc')}</option>
                <option value="custom" ${hasCustom ? 'selected' : ''}>${customLabel}</option>
            `;
        } else {
            return `
                <option value="top" ${order === 'top' ? 'selected' : ''}>${this._t('components.sortDialog.content.order.top')}</option>
                <option value="bottom" ${order === 'bottom' ? 'selected' : ''}>${this._t('components.sortDialog.content.order.bottom')}</option>
            `;
        }
    }

    /**
     * 更新次序选项
     */
    _updateOrderOptions(row, index) {
        const level = this.sortLevels[index];
        const orderSelect = row.querySelector('.sn-sort-order-select');
        orderSelect.innerHTML = this._getOrderOptions({ ...level, customOrder: null });

        if (level.sortBy === 'value') {
            orderSelect.value = 'asc';
            level.order = 'asc';
        } else {
            orderSelect.value = 'top';
            level.order = 'top';
        }
        level.customOrder = null;

        // 移除编辑按钮
        const editBtn = row.querySelector('.sn-sort-edit-custom');
        if (editBtn) editBtn.remove();
    }

    /**
     * 显示自定义列表弹窗
     */
    _showCustomListPopup(row, levelIndex) {
        // 移除已有的弹窗
        const existingPopup = this.overlay.querySelector('.sn-sort-custom-list-popup');
        if (existingPopup) existingPopup.remove();

        const level = this.sortLevels[levelIndex];
        const currentCustom = level.customOrder || [];

        // 计算弹窗位置（基于触发元素）
        const orderCol = row.querySelector('.sn-sort-col-order');
        const rect = orderCol.getBoundingClientRect();

        const popup = document.createElement('div');
        popup.className = 'sn-sort-custom-list-popup active';
        popup.innerHTML = `
            <div class="sn-sort-custom-list-header">
                <span>${this._t('components.sortDialog.content.customList.title')}</span>
                <button class="sn-sort-custom-list-close">×</button>
            </div>
            <div class="sn-sort-custom-list-presets">
                <label>${this._t('components.sortDialog.content.customList.presetLabel')}</label>
                <select class="sn-sort-preset-select">
                    <option value="">${this._t('components.sortDialog.content.customList.selectPreset')}</option>
                    ${this.customLists.map((list, i) =>
                        `<option value="${i}">${list.name}: ${list.values.slice(0, 3).join(', ')}...</option>`
                    ).join('')}
                </select>
            </div>
            <div class="sn-sort-custom-list-input">
                <label>${this._t('components.sortDialog.content.customList.customOrderLabel')}</label>
                <textarea class="sn-sort-custom-textarea" rows="6" placeholder="${this._t('components.sortDialog.content.customList.placeholder')}">${currentCustom.join('\n')}</textarea>
            </div>
            <div class="sn-sort-custom-list-footer">
                <button class="sn-sort-btn sn-sort-custom-cancel">${this._t('components.sortDialog.content.customList.cancel')}</button>
                <button class="sn-sort-btn sn-sort-custom-apply">${this._t('components.sortDialog.content.customList.apply')}</button>
            </div>
        `;

        // 挂载到 overlay 上，使用固定定位
        this.overlay.appendChild(popup);

        // 设置位置
        const popupWidth = 300;
        let left = rect.left;
        let top = rect.bottom + 4;

        // 确保不超出视口右边界
        if (left + popupWidth > window.innerWidth - 10) {
            left = window.innerWidth - popupWidth - 10;
        }
        // 确保不超出视口下边界
        if (top + 320 > window.innerHeight) {
            top = rect.top - 320 - 4;
        }

        popup.style.left = `${left}px`;
        popup.style.top = `${top}px`;

        // 关闭弹窗的辅助函数
        const closePopup = () => {
            popup.remove();
            if (!level.customOrder || level.customOrder.length === 0) {
                const orderSelect = row.querySelector('.sn-sort-order-select');
                if (orderSelect) {
                    orderSelect.value = 'asc';
                    level.order = 'asc';
                }
            }
        };

        // 事件绑定
        popup.querySelector('.sn-sort-custom-list-close').addEventListener('click', closePopup);

        popup.querySelector('.sn-sort-preset-select').addEventListener('change', (e) => {
            const idx = parseInt(e.target.value);
            if (!isNaN(idx) && this.customLists[idx]) {
                popup.querySelector('.sn-sort-custom-textarea').value = this.customLists[idx].values.join('\n');
            }
        });

        popup.querySelector('.sn-sort-custom-cancel').addEventListener('click', closePopup);

        popup.querySelector('.sn-sort-custom-apply').addEventListener('click', () => {
            const text = popup.querySelector('.sn-sort-custom-textarea').value.trim();
            const values = text.split('\n').map(s => s.trim()).filter(s => s);

            if (values.length === 0) {
                level.customOrder = null;
                level.order = 'asc';
            } else {
                level.customOrder = values;
                level.order = 'custom';
            }

            popup.remove();
            this._updateContent();
        });

        // 阻止事件冒泡
        popup.addEventListener('click', (e) => e.stopPropagation());
        popup.addEventListener('mousedown', (e) => e.stopPropagation());
    }

    /**
     * 选择条件行
     */
    _selectLevel(index) {
        this.selectedIndex = index;
        this._updateContent();
    }

    /**
     * 处理工具栏操作
     */
    _handleToolbarAction(action) {
        switch (action) {
            case 'add': this._addLevel(); break;
            case 'delete': this._deleteLevel(); break;
            case 'copy': this._copyLevel(); break;
            case 'moveUp': this._moveLevel(-1); break;
            case 'moveDown': this._moveLevel(1); break;
        }
    }

    /**
     * 添加条件
     */
    _addLevel() {
        const usedCols = new Set(this.sortLevels.map(l => l.column));
        let newCol = this.range.s.c;
        for (let c = this.range.s.c; c <= this.range.e.c; c++) {
            if (!usedCols.has(c)) {
                newCol = c;
                break;
            }
        }

        this.sortLevels.push({
            column: newCol,
            sortBy: 'value',
            order: 'asc',
            customOrder: null
        });

        this.selectedIndex = this.sortLevels.length - 1;
        this._updateContent();
    }

    /**
     * 删除条件
     */
    _deleteLevel() {
        if (this.sortLevels.length <= 1) return;
        if (this.selectedIndex === undefined) return;

        this.sortLevels.splice(this.selectedIndex, 1);
        if (this.selectedIndex >= this.sortLevels.length) {
            this.selectedIndex = this.sortLevels.length - 1;
        }
        this._updateContent();
    }

    /**
     * 复制条件
     */
    _copyLevel() {
        if (this.selectedIndex === undefined) return;

        const source = this.sortLevels[this.selectedIndex];
        const copy = { ...source, customOrder: source.customOrder ? [...source.customOrder] : null };
        this.sortLevels.splice(this.selectedIndex + 1, 0, copy);
        this.selectedIndex = this.selectedIndex + 1;
        this._updateContent();
    }

    /**
     * 移动条件
     */
    _moveLevel(direction) {
        if (this.selectedIndex === undefined) return;

        const newIndex = this.selectedIndex + direction;
        if (newIndex < 0 || newIndex >= this.sortLevels.length) return;

        const temp = this.sortLevels[this.selectedIndex];
        this.sortLevels[this.selectedIndex] = this.sortLevels[newIndex];
        this.sortLevels[newIndex] = temp;
        this.selectedIndex = newIndex;
        this._updateContent();
    }

    /**
     * 更新工具栏按钮状态
     */
    _updateToolbarState() {
        const hasSelection = this.selectedIndex !== undefined;
        const canDelete = this.sortLevels.length > 1;
        const canMoveUp = hasSelection && this.selectedIndex > 0;
        const canMoveDown = hasSelection && this.selectedIndex < this.sortLevels.length - 1;

        const toolbar = this.overlay.querySelector('.sn-sort-toolbar');
        toolbar.querySelector('[data-action="delete"]').disabled = !hasSelection || !canDelete;
        toolbar.querySelector('[data-action="copy"]').disabled = !hasSelection;
        toolbar.querySelector('[data-action="moveUp"]').disabled = !canMoveUp;
        toolbar.querySelector('[data-action="moveDown"]').disabled = !canMoveDown;
    }

    /**
     * 应用排序
     */
    _apply() {
        const sheet = this.SN.activeSheet;

        const sortKeys = this.sortLevels.map(level => ({
            col: level.column,
            order: level.order === 'desc' || level.order === 'bottom' ? 'desc' : 'asc',
            sortBy: level.sortBy,
            customOrder: level.customOrder
        }));

        const success = sortRange(sheet, this.range, sortKeys, {
            hasHeader: this.options.hasHeader,
            caseSensitive: this.options.caseSensitive
        });

        if (success) {
            this.SN._r();
            this.hide();
        }
    }

    /**
     * 显示对话框
     */
    _show() {
        this.selectedIndex = 0;
        this._updateContent();

        // 重置拖动位置
        const modal = this.overlay.querySelector('.sn-modal');
        if (modal) modal.style.transform = '';

        requestAnimationFrame(() => {
            this.overlay.classList.add('active');
        });
    }

    /**
     * 隐藏对话框
     */
    hide() {
        if (!this.overlay) return;

        this.overlay.classList.add('closing');
        setTimeout(() => {
            this.overlay.classList.remove('active', 'closing');
        }, 200);

        if (this._escHandler) {
            document.removeEventListener('keydown', this._escHandler);
        }
    }

    /**
     * 销毁对话框
     */
    destroy() {
        this.hide();
        if (this.overlay) {
            this.overlay.remove();
            this.overlay = null;
        }
    }
}
