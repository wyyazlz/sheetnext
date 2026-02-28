
import { initDropdown } from '../../components/dropdown.js';
import { initTooltip } from '../../components/tooltip.js';
import { setupViewport } from './helpers.js';

// 导入模块
import * as DOMBuilder from './DOMBuilder.js';
import * as ResizeHandler from './ResizeHandler.js';
import * as DragHandler from './DragHandler.js';
import * as EventManager from './EventManager.js';
import * as MenuConfig from './MenuConfig.js';
import { initToolsEvents, refreshActiveButtons as refreshActiveButtonsFn, refreshDisabledToolbarItems as refreshDisabledToolbarItemsFn, syncToolbarState } from './ToolbarBuilder.js';
import getSvg from '../../assets/mainSvgs.js';
import PivotDataField from '../PivotTable/PivotDataField.js';
import StateSync from './StateSync.js';

/**
 * 布局与工具栏管理
 * @title 🎨 UI布局
 * @class
 */
export default class Layout {

    /**
     * @param {Object} SN - SheetNext instance.
     * @param {Object} [options={}] - Layout options.
     * @param {function} [options.menuRight] - Callback `(defaultHTML: string) => string` to customize the top-right menu HTML.
     * @param {function} [options.menuList] - Callback `(config: Array) => Array` to customize menu tab config.
     */
    constructor(SN, options = {}) {
        this._options = options;
        /**
         * SheetNext 主实例
         * @type {Object}
         */
        this.SN = SN;
        this._isInitialized = false;

        // 状态同步器（工具栏 checkbox 与 API 状态双向绑定）
        /**
         * 状态同步器
         * @type {StateSync}
         */
        this.StateSync = new StateSync(SN);

        // 状态属性
        this._showMenuBar = true;
        this._showToolbar = true;
        this._showAIChat = false;
        this._showFormulaBar = true;
        this._showSheetTabBar = true;
        this._showStats = true;
        this._showAIChatWindow = false;
        this._showPivotPanel = false;
        this._activePivotTable = null;  // 当前编辑的透视表
        this._pivotPanelManualClosed = false;  // 用户是否手动关闭了面板        
        this._currentContextType = null;
        this._currentContextData = null;
        this._toolbarMode = 'pro';
        this._userToolbarMode = 'pro';
        this._toolbarModeLocked = false;
        this._toolbarAdaptiveRaf = null;
        this._minimalToolbarEventsBound = false;
        this._minimalToolbarTextEnabled = false;

        // UI组件
        /**
         * 通用弹窗实例
         * @type {Object|null}
         */
        this.myModal = null;
        /**
         * Toast 实例
         * @type {Object|null}
         */
        this.toast = null;

        // 拖拽相关
        /**
         * 聊天窗口拖拽偏移 X
         * @type {number}
         */
        this.chatWindowOffsetX = 0;
        /**
         * 聊天窗口拖拽偏移 Y
         * @type {number}
         */
        this.chatWindowOffsetY = 0;
        this._lastIsSmallWindow = null;
        this._resizeObserver = null;
        this._resizeTimer = null;
        this._lastContainerWidth = null;
        this._lastContainerHeight = null;

        // 菜单配置
        /**
         * 菜单配置
         * @type {Object}
         */
        this.menuConfig = MenuConfig.initDefaultMenuConfig(SN, options);

        // 一次性初始化所有内容
        this._init(SN.containerDom);
    }

    // ============================= Getters & Setters =============================

    /**
     * 是否为小窗口
     * @type {boolean}     */
    get isSmallWindow() {
        return this.SN.containerDom.offsetWidth < 900;
    }

    get toolbarMode() {
        return this._toolbarMode;
    }

    get isToolbarModeLocked() {
        return this._toolbarModeLocked;
    }

    get minimalToolbarTextEnabled() {
        return this._minimalToolbarTextEnabled;
    }

    /**
     * 是否显示菜单栏
     * @type {boolean}
     */
    get showMenuBar() {
        return this._showMenuBar;
    }

    set showMenuBar(bol) {
        this._showMenuBar = bol;
        if (bol) {
            this.snHead.style.display = "flex";
        } else {
            this.snHead.style.display = "none";
        }
        this.updateCanvasSize();
    }

    /**
     * 是否显示工具栏
     * @type {boolean}
     */
    get showToolbar() {
        return this._showToolbar;
    }

    set showToolbar(bol) {
        this._showToolbar = bol;
        if (bol) {
            this.snTools.style.display = "flex";
        } else {
            this.snTools.style.display = "none";
        }
        this.updateCanvasSize();
    }

    /**
     * 是否显示 AI 聊天入口
     * @type {boolean}
     */
    get showAIChat() {
        return this._showAIChat;
    }

    set showAIChat(bol) {
        this._showAIChat = bol;
        // 通知工具栏 checkbox 同步状态
        this.StateSync.notify('showAIChat', bol);
        if (bol) {
            this.snChat.style.display = "block";

            if (this.isSmallWindow) {
                this.showAIChatWindow = true;
            }
        } else {
            this.snChat.style.display = "none";
        }
        this.updateCanvasSize();
    }

    /**
     * 是否显示公式栏
     * @type {boolean}
     */
    get showFormulaBar() {
        return this._showFormulaBar;
    }

    set showFormulaBar(bol) {
        this._showFormulaBar = bol;
        // 通知工具栏 checkbox 同步状态
        this.StateSync.notify('showFormulaBar', bol);
        if (bol) {
            this.snFormulaBar.style.display = "grid";
        } else {
            this.snFormulaBar.style.display = "none";
        }
        this.updateCanvasSize();
    }

    /**
     * 是否显示工作表标签栏
     * @type {boolean}
     */
    get showSheetTabBar() {
        return this._showSheetTabBar;
    }

    set showSheetTabBar(bol) {
        this._showSheetTabBar = bol;
        // 通知工具栏 checkbox 同步状态
        this.StateSync.notify('showSheetTabBar', bol);
        // Use CSS default (flex) instead of forcing grid.
        this.snSheetTabBar.style.display = bol ? "" : "none";
        this.updateCanvasSize();
    }

    /**
     * 是否显示统计栏
     * @type {boolean}
     */
    get showStats() {
        return this._showStats;
    }

    set showStats(bol) {
        this._showStats = bol;
        // 通知工具栏 checkbox 同步状态
        this.StateSync.notify('showStats', bol);
        const statsItems = this.SN.containerDom.querySelectorAll('.sn-stat-item');
        statsItems.forEach(item => {
            item.style.display = bol ? '' : 'none';
        });
    }

    /**
     * 是否显示 AI 聊天窗口
     * @type {boolean}
     */
    get showAIChatWindow() {
        return this._showAIChatWindow;
    }

    set showAIChatWindow(bol) {
        if (!bol && this.isSmallWindow && this._showAIChat) {
            // 在小屏幕下点击收起按钮时，将侧边栏设为false不显示
            this.showAIChat = false;
        }

        this._showAIChatWindow = bol;
        if (bol) { // 显示小窗
            this.snChat.classList.add("sn-chat-window");
            this.snChat.style.transform = `translate(${this.chatWindowOffsetX}px, ${this.chatWindowOffsetY}px)`;
        } else {
            this.snChat.classList.remove("sn-chat-window");
            this.snChat.style.transform = "none";
        }

        this.updateCanvasSize();
    }

    /**
     * 是否显示透视表字段面板
     * @type {boolean}
     */
    get showPivotPanel() {
        return this._showPivotPanel;
    }

    set showPivotPanel(bol) {
        const prevPivotTable = this._activePivotTable;
        this._showPivotPanel = bol;
        if (bol) {
            this.snPivotPanel.style.display = 'flex';
        } else {
            this.snPivotPanel.style.display = 'none';
            this._activePivotTable = null;
        }
        const updateTarget = bol ? this._activePivotTable : prevPivotTable;
        if (updateTarget) {
            this._updatePivotAnalyzePanel(updateTarget);
        }
        this.updateCanvasSize();
    }

    /**
     * 打开透视表字段面板（用户手动调用，如工具栏按钮）
     * @param {PivotTable} pt - 透视表实例
     * @returns {void}
     */
    openPivotPanel(pt) {
        this._pivotPanelManualClosed = false;  // 重置手动关闭标记
        this._activePivotTable = pt;
        this._renderPivotPanel(pt);
        this.showPivotPanel = true;
    }

    /**
     * 切换透视表字段面板显示
     * @param {PivotTable} pt - 透视表实例
     * @returns {boolean} 是否显示
     */
    togglePivotPanel(pt) {
        if (this._showPivotPanel && this._activePivotTable === pt) {
            this._pivotPanelManualClosed = true;  // 标记为手动关闭
            this.showPivotPanel = false;
            return false;
        }
        this.openPivotPanel(pt);
        return true;
    }

    /**
     * 自动打开透视表字段面板（点击透视表区域时调用）
     * 如果用户手动关闭过，则不自动打开
     * @param {PivotTable} pt - 透视表实例
     * @returns {void}
     */
    autoOpenPivotPanel(pt) {
        if (this._pivotPanelManualClosed) return;  // 用户手动关闭过，不自动打开
        if (this._showPivotPanel && this._activePivotTable === pt) return;  // 已打开同一透视表，避免重复
        this._activePivotTable = pt;
        this._renderPivotPanel(pt);
        this.showPivotPanel = true;
    }

    /**
     * 关闭透视表字段面板（点击透视表区域外时调用）
     * 不会设置手动关闭标记
     * @returns {void}
     */
    closePivotPanel() {
        if (!this._showPivotPanel) return;  // 已经关闭，避免循环调用
        this.showPivotPanel = false;
    }

    /**
     * 刷新工具栏中带有 active getter 的按钮状态
     * 用于格式刷等切换按钮的双向绑定
     * @returns {void}
     */
    refreshActiveButtons() {
        if (this._toolbarBuilder) {
            refreshActiveButtonsFn(this.SN.containerDom, this._toolbarBuilder);
        }
    }

    /**
     * 统一刷新工具栏状态（面板切换、API修改后调用）
     * 整合：样式状态 + checkbox + active按钮
     * @returns {void}
     */
    refreshToolbar() {
        const sheet = this.SN.activeSheet;
        if (!sheet) return;

        const { r, c } = sheet.activeCell;
        const cell = sheet.getCell(r, c);

        // 1. sync toolbar style state (font/size/alignment/etc.)
        syncToolbarState(this.SN.containerDom, cell, this._toolbarBuilder?.menuConfig);

        // 2. sync active-toggle buttons (e.g. format brush)
        this.refreshActiveButtons();

        // 3. sync checkbox state in active panel
        const activePanel = this.SN.containerDom.querySelector('.sn-tools-panel.active');
        this._refreshPanelCheckboxStates(activePanel);

        // 4. sync disabled state in active panel
        refreshDisabledToolbarItemsFn(this.SN.containerDom, this._toolbarBuilder);

        this._scheduleToolbarAdaptiveLayout();
    }

    _refreshPanelCheckboxStates(panel) {
        if (!panel || !this._toolbarBuilder) return;
        const checkboxes = panel.querySelectorAll('input[type="checkbox"]');
        const config = this._toolbarBuilder.config;

        checkboxes.forEach(checkbox => {
            const fieldId = checkbox.dataset.field;
            for (const panelConfig of config) {
                for (const group of panelConfig.groups || []) {
                    for (const item of group.items || []) {
                        const found = this._findCheckboxInItem(item, fieldId);
                        if (!found) continue;

                        if (typeof found.checked === 'function') {
                            checkbox.checked = !!found.checked(this.SN);
                        }

                        if (typeof found.disabled === 'function') {
                            checkbox.disabled = !!found.disabled(this.SN);
                        } else if (typeof found.disabled !== 'undefined') {
                            checkbox.disabled = !!found.disabled;
                        }
                        return;
                    }
                }
            }
        });
    }
    _findCheckboxInItem(item, fieldId) {
        if (item.type === 'checkbox' && item.id === fieldId) return item;
        if ((item.type === 'stack' || item.type === 'row') && Array.isArray(item.items)) {
            for (const child of item.items) {
                const found = this._findCheckboxInItem(child, fieldId);
                if (found) return found;
            }
        }
        return null;
    }

    toggleMinimalToolbar(enabled) {
        const nextEnabled = !!enabled;
        if (!nextEnabled && this._toolbarModeLocked) {
            this._syncMinimalToolbarCheckboxUI();
            return;
        }

        this._userToolbarMode = nextEnabled ? 'minimal' : 'pro';
        this._setToolbarMode(this._userToolbarMode);
        this._syncMinimalToolbarCheckboxUI();
    }

    toggleMinimalToolbarText(enabled) {
        this._setMinimalToolbarTextEnabled(enabled);
        this._syncMinimalToolbarCheckboxUI();
        this._scheduleToolbarAdaptiveLayout();
    }

    _setMinimalToolbarTextEnabled(enabled) {
        const nextEnabled = !!enabled;
        this._minimalToolbarTextEnabled = nextEnabled;
        this.SN.containerDom.classList.toggle('sn-toolbar-minimal-text', nextEnabled);
    }

    _syncToolbarModeByWidth(width = this.SN.containerDom.offsetWidth) {
        const shouldLock = width < 1400;
        this._toolbarModeLocked = shouldLock;
        const nextMode = shouldLock ? 'minimal' : this._userToolbarMode;
        this._setToolbarMode(nextMode);
        this._syncMinimalToolbarCheckboxUI();
    }

    _setToolbarMode(mode) {
        const normalized = mode === 'minimal' ? 'minimal' : 'pro';
        if (this._toolbarMode === normalized) {
            this._scheduleToolbarAdaptiveLayout();
            return;
        }

        this._toolbarMode = normalized;
        this.SN.containerDom.classList.toggle('sn-toolbar-minimal', normalized === 'minimal');
        this._scheduleToolbarAdaptiveLayout();
        requestAnimationFrame(() => this.updateCanvasSize());
    }

    _syncMinimalToolbarCheckboxUI() {
        this._syncMenuMinimalToggleUI();
        const activePanel = this.SN.containerDom.querySelector('.sn-tools-panel.active');
        if (!activePanel || !this._toolbarBuilder) return;
        this._refreshPanelCheckboxStates(activePanel);
        refreshDisabledToolbarItemsFn(this.SN.containerDom, this._toolbarBuilder);
    }

    _syncMenuMinimalToggleUI() {
        const menuToggle = this.SN.containerDom.querySelector('.sn-menu-minimal-toggle input[data-role="menu-minimal-toolbar-toggle"]');
        if (!menuToggle) return;
        const isMinimal = this._toolbarMode === 'minimal';
        menuToggle.checked = isMinimal;
        menuToggle.disabled = this._toolbarModeLocked && isMinimal;
    }

    _bindMenuRightEvents() {
        const menuToggle = this.SN.containerDom.querySelector('.sn-menu-minimal-toggle input[data-role="menu-minimal-toolbar-toggle"]');
        if (!menuToggle) return;
        menuToggle.addEventListener('change', () => {
            this.toggleMinimalToolbar(menuToggle.checked);
        });
    }

    _bindMinimalToolbarEvents() {
        if (this._minimalToolbarEventsBound) return;
        this._minimalToolbarEventsBound = true;

        const menuList = this.SN.containerDom.querySelector('.sn-menu-list');
        if (!menuList) return;

        menuList.addEventListener('click', (e) => {
            const item = e.target.closest('.sn-tab-overflow-menu li[data-panel]');
            if (!item) return;

            const panelName = item.dataset.panel;
            const tab = menuList.querySelector('.sn-menu-tab[data-panel="' + panelName + '"]');
            if (tab) this._switchToTab(tab);
        });

        this.SN.containerDom.addEventListener('mousedown', (e) => {
            const host = e.target.closest('.sn-toolbar-more');
            if (!host || !this.SN.containerDom.contains(host)) return;
            // 点击溢出菜单内部不处理
            if (e.target.closest('.sn-toolbar-overflow-menu')) return;

            e.preventDefault();
            e.stopPropagation();

            const wasActive = host.classList.contains('active');
            this.SN.containerDom.querySelectorAll('.sn-toolbar-more.active').forEach(el => el.classList.remove('active'));
            if (!wasActive) {
                const menu = host.querySelector('.sn-toolbar-overflow-menu');
                if (menu) {
                    const rect = host.getBoundingClientRect();
                    const containerRect = this.SN.containerDom.getBoundingClientRect();
                    menu.style.top = rect.bottom + 10 + 'px';
                    menu.style.right = (document.documentElement.clientWidth - containerRect.right + 8) + 'px';
                    menu.style.left = '';
                }
                host.classList.add('active');
            }
        });

        window.addEventListener('mousedown', (e) => {
            this.SN.containerDom.querySelectorAll('.sn-toolbar-more.active').forEach(host => {
                if (!host.contains(e.target)) host.classList.remove('active');
            });
        });
    }

    _scheduleToolbarAdaptiveLayout() {
        if (this._toolbarAdaptiveRaf !== null) {
            cancelAnimationFrame(this._toolbarAdaptiveRaf);
        }

        this._toolbarAdaptiveRaf = requestAnimationFrame(() => {
            this._toolbarAdaptiveRaf = null;
            this._applyToolbarAdaptiveLayout();
        });
    }

    _applyToolbarAdaptiveLayout() {
        if (this._toolbarMode !== 'minimal') {
            this._restoreTabOverflow();
            this._restoreAllPanelOverflow();
            return;
        }

        this._applyMinimalTabOverflow();
        this._applyMinimalPanelOverflow();
    }

    _ensureTabOverflowHost(menuList) {
        let host = menuList.querySelector('.sn-tab-overflow');
        if (host) return host;

        host = document.createElement('div');
        host.className = 'sn-dropdown sn-tab-overflow';

        const toggle = document.createElement('div');
        toggle.className = 'sn-dropdown-toggle sn-no-icon';
        const title = document.createElement('span');
        title.textContent = '\u66f4\u591a';
        const arrow = document.createElement('i');
        arrow.className = 'sn-arrow';
        toggle.appendChild(title);
        toggle.appendChild(arrow);

        const menu = document.createElement('ul');
        menu.className = 'sn-dropdown-menu sn-tab-overflow-menu';

        host.appendChild(toggle);
        host.appendChild(menu);
        menuList.appendChild(host);
        return host;
    }

    _restoreTabOverflow() {
        const menuList = this.SN.containerDom.querySelector('.sn-menu-list');
        if (!menuList) return;

        menuList.querySelectorAll('.sn-menu-tab[data-minimal-overflow-hidden="true"]').forEach(tab => {
            const shouldHideByContext = tab.classList.contains('sn-contextual-tab')
                && tab.dataset.contextType !== (this._currentContextType || '');
            tab.style.display = shouldHideByContext ? 'none' : '';
            delete tab.dataset.minimalOverflowHidden;
        });

        const host = menuList.querySelector('.sn-tab-overflow');
        if (!host) return;

        const menu = host.querySelector('.sn-tab-overflow-menu');
        if (menu) menu.innerHTML = '';
        host.style.display = 'none';
    }

    _applyMinimalTabOverflow() {
        const menuList = this.SN.containerDom.querySelector('.sn-menu-list');
        if (!menuList) return;

        this._restoreTabOverflow();

        const visibleTabs = Array.from(menuList.querySelectorAll('.sn-menu-tab')).filter(tab => getComputedStyle(tab).display !== 'none');
        if (!visibleTabs.length) return;

        const host = this._ensureTabOverflowHost(menuList);
        host.style.display = 'flex';
        const overflowMenu = host.querySelector('.sn-tab-overflow-menu');
        overflowMenu.innerHTML = '';

        const orderMap = new Map(visibleTabs.map((tab, index) => [tab, index]));
        const activeTab = menuList.querySelector('.sn-menu-tab.active');
        const availableWidth = Math.max(0, menuList.clientWidth - host.offsetWidth - 6);

        let used = 0;
        const kept = [];
        let hidden = [];

        visibleTabs.forEach(tab => {
            const tabWidth = tab.offsetWidth;
            if (used + tabWidth <= availableWidth) {
                kept.push(tab);
                used += tabWidth;
            } else {
                hidden.push(tab);
            }
        });

        if (activeTab && hidden.includes(activeTab)) {
            hidden = hidden.filter(tab => tab !== activeTab);
            const fallback = [...kept].reverse().find(tab => tab !== activeTab);
            if (fallback) {
                kept.splice(kept.indexOf(fallback), 1);
                hidden.unshift(fallback);
            }
            kept.push(activeTab);
        }

        hidden.sort((a, b) => (orderMap.get(a) || 0) - (orderMap.get(b) || 0));

        hidden.forEach(tab => {
            tab.style.display = 'none';
            tab.dataset.minimalOverflowHidden = 'true';

            const item = document.createElement('li');
            item.dataset.panel = tab.dataset.panel || '';
            item.textContent = (tab.textContent || '').trim();
            if (tab.classList.contains('active')) item.classList.add('active');
            overflowMenu.appendChild(item);
        });

        if (!hidden.length) {
            host.style.display = 'none';
        }
    }

    _ensurePanelOverflowHost(panel) {
        let host = Array.from(panel.children).find(child => child.classList?.contains('sn-toolbar-more'));
        if (host) return host;

        host = document.createElement('div');
        host.className = 'sn-tools-item small sn-toolbar-more';

        const toggle = document.createElement('div');
        toggle.className = 'sn-toolbar-more-toggle';
        toggle.innerHTML = getSvg('zk');

        const menu = document.createElement('div');
        menu.className = 'sn-toolbar-overflow-menu';

        host.appendChild(toggle);
        host.appendChild(menu);
        panel.appendChild(host);
        return host;
    }

    _restorePanelOverflow(panel) {
        const host = Array.from(panel.children).find(child => child.classList?.contains('sn-toolbar-more'));
        if (!host) return;

        const overflowMenu = host.querySelector('.sn-toolbar-overflow-menu');
        const movedGroups = Array.from(overflowMenu.children).filter(child => child.classList?.contains('sn-tools-group'));
        movedGroups.sort((a, b) => Number(a.dataset.minimalOrder || 0) - Number(b.dataset.minimalOrder || 0));
        movedGroups.forEach(group => panel.insertBefore(group, host));

        overflowMenu.innerHTML = '';
        host.classList.remove('active');
        host.style.display = 'none';
    }

    _restoreAllPanelOverflow() {
        const panels = this.SN.containerDom.querySelectorAll('.sn-tools-panel');
        panels.forEach(panel => this._restorePanelOverflow(panel));
    }

    _applyMinimalPanelOverflow() {
        const activePanel = this.SN.containerDom.querySelector('.sn-tools-panel.active');
        if (!activePanel) return;

        if (this._toolbarBuilder && this._toolbarBuilder.renderPanelIfNeeded(activePanel)) {
            this.initColorComponents(activePanel);
        }

        this._restorePanelOverflow(activePanel);

        const groups = Array.from(activePanel.children).filter(child => child.classList?.contains('sn-tools-group'));
        if (!groups.length) return;

        groups.forEach((group, index) => {
            if (!group.dataset.minimalOrder) group.dataset.minimalOrder = String(index);
        });

        const host = this._ensurePanelOverflowHost(activePanel);
        const overflowMenu = host.querySelector('.sn-toolbar-overflow-menu');
        overflowMenu.innerHTML = '';

        host.style.display = 'flex';
        const availableWidth = Math.max(0, activePanel.clientWidth - host.offsetWidth - 6);

        let used = 0;
        const hiddenGroups = [];
        groups.forEach((group, index) => {
            const width = group.offsetWidth;
            if (index === 0 || used + width <= availableWidth) {
                used += width;
            } else {
                hiddenGroups.push(group);
            }
        });

        if (!hiddenGroups.length) {
            host.classList.remove('active');
            host.style.display = 'none';
            return;
        }

        hiddenGroups.forEach(group => overflowMenu.appendChild(group));
    }

    _renderPivotPanel(pt) {
        const cache = pt.cache;
        if (!cache || !cache.fields.length) {
            this.SN.Utils.toast(this.SN.t('core.layout.layout.toast.msg001'));
            return;
        }

        const panel = this.snPivotPanel;
        const ns = this.SN.namespace;

        // 渲染字段列表
        const fieldsList = panel.querySelector('.sn-pivot-fields-list');
        fieldsList.innerHTML = cache.fields.map((field, index) => `
            <div class="sn-pivot-field-item" data-index="${index}" draggable="true">
                <label class="sn-pivot-field-label">
                    <input type="checkbox" data-index="${index}">
                    <span class="sn-pivot-field-text">${field.name}</span>
                </label>
            </div>
        `).join('');

        // 清空各区域
        panel.querySelectorAll('.sn-pivot-zone-list').forEach(zone => {
            zone.innerHTML = '';
        });

        // 同步当前透视表配置到面板
        this._syncPivotPanelState(pt);

        // 绑定事件
        this._bindPivotPanelEvents(pt);
    }

    _syncPivotPanelState(pt) {
        const panel = this.snPivotPanel;
        const cache = pt.cache;
        const getFieldName = (idx) => pt.pivotFields[idx]?.name || cache.fields[idx]?.name || this.SN.t('core.layout.layout.pivotPanel.fieldFallback', { idx });

        // 更新行字段区域
        const rowZone = panel.querySelector('.sn-pivot-zone[data-zone="row"] .sn-pivot-zone-list');
        rowZone.innerHTML = pt.rowFields.map(idx => this._createZoneItem(idx, getFieldName(idx))).join('');

        // 更新列字段区域
        const colZone = panel.querySelector('.sn-pivot-zone[data-zone="column"] .sn-pivot-zone-list');
        colZone.innerHTML = pt.colFields.map(idx => this._createZoneItem(idx, getFieldName(idx))).join('');

        // 更新值字段区域
        const valueZone = panel.querySelector('.sn-pivot-zone[data-zone="value"] .sn-pivot-zone-list');
        valueZone.innerHTML = pt.dataFields.map((df, dfIdx) => {
            const label = df.name || getFieldName(df.fld);
            return this._createZoneItem(df.fld, label, df.subtotal, dfIdx);
        }).join('');

        // 更新筛选字段区域
        const filterZone = panel.querySelector('.sn-pivot-zone[data-zone="filter"] .sn-pivot-zone-list');
        filterZone.innerHTML = pt.pageFields.map(idx => this._createZoneItem(idx, getFieldName(idx))).join('');

        // 更新字段列表的勾选状态
        const usedIndices = new Set([
            ...pt.rowFields,
            ...pt.colFields,
            ...pt.dataFields.map(df => df.fld),
            ...pt.pageFields
        ]);

        panel.querySelectorAll('.sn-pivot-field-item input[type="checkbox"]').forEach(checkbox => {
            checkbox.checked = usedIndices.has(Number(checkbox.dataset.index));
        });
    }

    _createZoneItem(index, name, subtotal = '', dataFieldIndex = '') {
        const t = (key, params) => this.SN.t(key, params);
        return `
            <div class="sn-pivot-zone-item" data-index="${index}" data-subtotal="${subtotal}" data-df-index="${dataFieldIndex}" draggable="true">
                <span class="sn-pivot-zone-item-name">${name}</span>
                <div class="sn-pivot-zone-item-actions">
                    <button class="sn-pivot-zone-item-action sn-pivot-zone-item-remove" title="${t('core.layout.layout.pivotPanel.remove')}">✕</button>
                    <button class="sn-pivot-zone-item-action sn-pivot-zone-item-settings" title="${t('core.layout.layout.pivotPanel.fieldSettings')}">⚙</button>
                </div>
            </div>
        `;
    }

    _bindPivotPanelEvents(pt) {
        const panel = this.snPivotPanel;
        const SN = this.SN;

        // 关闭按钮
        panel.querySelector('[data-action="close"]').onclick = () => {
            this._pivotPanelManualClosed = true;  // 标记为手动关闭
            this.showPivotPanel = false;
        };

        // 搜索功能
        const searchInput = panel.querySelector('.sn-pivot-search-input');
        searchInput.oninput = () => {
            const keyword = searchInput.value.toLowerCase();
            panel.querySelectorAll('.sn-pivot-field-item').forEach(item => {
                const label = item.querySelector('label').textContent.toLowerCase();
                item.style.display = label.includes(keyword) ? '' : 'none';
            });
        };

        // 字段勾选事件
        panel.querySelectorAll('.sn-pivot-field-item input[type="checkbox"]').forEach(checkbox => {
            checkbox.onchange = () => {
                const index = Number(checkbox.dataset.index);
                if (checkbox.checked) {
                    const zone = this._getDefaultFieldZone(pt, index);
                    this._addFieldToZone(pt, zone, index, 'sum');
                } else {
                    // 从所有区域移除
                    this._removeFieldFromAllZones(pt, index);
                }
                this._syncPivotPanelState(pt);
                this._updatePivotTableIfNeeded(pt);
            };
        });

        // 区域内移除按钮
        panel.querySelectorAll('.sn-pivot-zone-list').forEach(zoneList => {
            zoneList.onclick = (e) => {
                const item = e.target.closest('.sn-pivot-zone-item');
                if (!item) return;
                const zone = zoneList.closest('.sn-pivot-zone').dataset.zone;
                const index = Number(item.dataset.index);
                const dfIndex = item.dataset.dfIndex !== '' ? Number(item.dataset.dfIndex) : null;

                if (e.target.closest('.sn-pivot-zone-item-settings')) {
                    this._openPivotFieldSettingsModal(pt, zone, index, dfIndex);
                    return;
                }

                if (e.target.closest('.sn-pivot-zone-item-remove')) {
                    this._removeFieldFromZone(pt, zone, index, dfIndex);
                    this._syncPivotPanelState(pt);
                    this._updatePivotTableIfNeeded(pt);
                }
            };
        });

        // 更新按钮
        panel.querySelector('.sn-pivot-update-btn').onclick = () => {
            pt.calculate();
            pt.render();
            SN.Utils.toast(SN.t('core.layout.layout.toast.msg002'));
        };

        // 延迟更新checkbox
        const delayCheckbox = panel.querySelector('.sn-pivot-delay-update input');
        this._delayUpdate = delayCheckbox.checked;
        delayCheckbox.onchange = () => {
            this._delayUpdate = delayCheckbox.checked;
        };

        // 绑定拖拽事件
        this._bindPivotDragEvents(pt);
    }

    _bindPivotDragEvents(pt) {
        const panel = this.snPivotPanel;
        let dragData = null;

        // 字段列表拖拽开始
        panel.querySelectorAll('.sn-pivot-field-item').forEach(item => {
            item.ondragstart = (e) => {
                const index = Number(item.dataset.index);
                const name = item.querySelector('label').textContent;
                dragData = { type: 'field', index, name, sourceZone: null };
                item.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', index);
            };

            item.ondragend = () => {
                item.classList.remove('dragging');
                dragData = null;
            };
        });

        // 区域拖拽事件
        panel.querySelectorAll('.sn-pivot-zone').forEach(zone => {
            const zoneType = zone.dataset.zone;
            const zoneList = zone.querySelector('.sn-pivot-zone-list');

            // 区域内项目拖拽开始
            zoneList.ondragstart = (e) => {
                const item = e.target.closest('.sn-pivot-zone-item');
                if (!item) return;

                const index = Number(item.dataset.index);
                const name = item.querySelector('.sn-pivot-zone-item-name').textContent;
                const subtotal = item.dataset.subtotal || '';
                const dfIndex = item.dataset.dfIndex !== '' ? Number(item.dataset.dfIndex) : null;
                dragData = { type: 'zone-item', index, name, subtotal, dfIndex, sourceZone: zoneType };
                item.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', index);
            };

            zoneList.ondragend = (e) => {
                const item = e.target.closest('.sn-pivot-zone-item');
                if (item) item.classList.remove('dragging');
                dragData = null;
            };

            // 拖拽经过区域
            zone.ondragover = (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                zone.classList.add('drag-over');
            };

            zone.ondragleave = (e) => {
                // 只在真正离开区域时移除样式
                if (!zone.contains(e.relatedTarget)) {
                    zone.classList.remove('drag-over');
                }
            };

            // 放置到区域
            zone.ondrop = (e) => {
                e.preventDefault();
                zone.classList.remove('drag-over');

                if (!dragData) return;

                const targetZone = zoneType;
                const { index, sourceZone, subtotal, dfIndex } = dragData;
                const dropIndex = this._getPivotZoneDropIndex(zoneList, e);

                if (sourceZone === targetZone) {
                    if (targetZone === 'value') {
                        if (typeof dfIndex === 'number' && !Number.isNaN(dfIndex)) {
                            const target = dropIndex > dfIndex ? dropIndex - 1 : dropIndex;
                            pt.moveDataField(dfIndex, target);
                        }
                    } else {
                        this._moveFieldWithinZone(pt, targetZone, index, dropIndex);
                    }
                } else {
                    // 如果从其他区域拖来，先从源区域移除
                    if (sourceZone) {
                        this._removeFieldFromZone(pt, sourceZone, index, dfIndex);
                    }
                    // 添加到目标区域
                    this._addFieldToZone(pt, targetZone, index, subtotal, dropIndex);
                }

                this._syncPivotPanelState(pt);
                this._updatePivotTableIfNeeded(pt);
                dragData = null;
            };
        });

        const fieldsList = panel.querySelector('.sn-pivot-fields-list');
        if (fieldsList) {
            fieldsList.ondragover = (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                fieldsList.classList.add('drag-over');
            };

            fieldsList.ondragleave = (e) => {
                if (!fieldsList.contains(e.relatedTarget)) {
                    fieldsList.classList.remove('drag-over');
                }
            };

            fieldsList.ondrop = (e) => {
                e.preventDefault();
                fieldsList.classList.remove('drag-over');
                if (!dragData || dragData.type !== 'zone-item' || !dragData.sourceZone) return;

                const { index, sourceZone, dfIndex } = dragData;
                this._removeFieldFromZone(pt, sourceZone, index, dfIndex);
                this._syncPivotPanelState(pt);
                this._updatePivotTableIfNeeded(pt);
                dragData = null;
            };
        }
    }

    _openPivotFieldSettingsModal(pt, zone, fieldIndex, dataFieldIndex) {
        if (zone === 'value') {
            if (typeof dataFieldIndex !== 'number' || Number.isNaN(dataFieldIndex)) return;
            this._openPivotValueFieldSettingsModal(pt, dataFieldIndex);
            return;
        }
        this._openPivotAxisFieldSettingsModal(pt, fieldIndex);
    }

    _openPivotAxisFieldSettingsModal(pt, fieldIndex) {
        const SN = this.SN;
        const field = pt.pivotFields[fieldIndex];
        if (!field) return;
        const cacheField = pt.cache?.fields?.[fieldIndex];
        const sourceName = cacheField?.name || field.name || `Field${fieldIndex}`;
        const customName = field.name || sourceName;
        const subtotalMode = field.defaultSubtotal === false ? 'none' : 'auto';
        const subtotalPosition = field.subtotalTop === false ? 'bottom' : 'top';
        const sortType = field.sortType || 'manual';
        const showDropDowns = field.showDropDowns !== false;
        const showAll = field.showAll === true;
        const repeatLabels = field.repeatItemLabels === true;
        const insertBlank = field.insertBlankRow === true;
        const layoutMode = pt.compact ? 'compact' : (pt.outline ? 'outline' : 'tabular');

        SN.Utils.modal({
            titleKey: 'core.layout.layout.modal.m001.title', title: 'Field Settings',
            maxWidth: '480px',
            content: `
                <div class="sn-pivot-settings">
                    <div class="sn-pivot-settings-row">
                        <label>Source Name</label>
                        <div class="sn-pivot-settings-static">${sourceName}</div>
                    </div>
                    <div class="sn-pivot-settings-row">
                        <label>Custom Name</label>
                        <input type="text" data-field="pivot-custom-name" value="${customName}">
                    </div>
                    <div class="sn-pivot-settings-tabs">
                        <button type="button" class="sn-pivot-settings-tab active" data-tab="summary">Subtotal & Filter</button>
                        <button type="button" class="sn-pivot-settings-tab" data-tab="layout">Layout & Print</button>
                    </div>
                    <div class="sn-pivot-settings-panel active" data-panel="summary">
                        <div class="sn-pivot-settings-section">
                            <div class="sn-pivot-settings-section-title">Subtotal</div>
                            <label class="sn-pivot-settings-option">
                                <input type="radio" name="pivot-subtotal" value="auto"${subtotalMode === 'auto' ? ' checked' : ''}>
                                <span>Automatic</span>
                            </label>
                            <label class="sn-pivot-settings-option">
                                <input type="radio" name="pivot-subtotal" value="none"${subtotalMode === 'none' ? ' checked' : ''}>
                                <span>None</span>
                            </label>
                            <div class="sn-pivot-settings-row">
                                <label>Position</label>
                                <select data-field="pivot-subtotal-position">
                                    <option value="top"${subtotalPosition === 'top' ? ' selected' : ''}>Top</option>
                                    <option value="bottom"${subtotalPosition === 'bottom' ? ' selected' : ''}>Bottom</option>
                                </select>
                            </div>
                        </div>
                        <div class="sn-pivot-settings-section">
                            <div class="sn-pivot-settings-section-title">Filter</div>
                            <label class="sn-pivot-settings-option">
                                <input type="checkbox" data-field="pivot-show-dropdowns"${showDropDowns ? ' checked' : ''}>
                                <span>Show Filter Button</span>
                            </label>
                            <label class="sn-pivot-settings-option">
                                <input type="checkbox" data-field="pivot-show-all"${showAll ? ' checked' : ''}>
                                <span>Include New Items</span>
                            </label>
                            <div class="sn-pivot-settings-row">
                                <label>Sort</label>
                                <select data-field="pivot-sort-type">
                                    <option value="manual"${sortType === 'manual' ? ' selected' : ''}>Manual</option>
                                    <option value="ascending"${sortType === 'ascending' ? ' selected' : ''}>Ascending</option>
                                    <option value="descending"${sortType === 'descending' ? ' selected' : ''}>Descending</option>
                                </select>
                            </div>
                        </div>
                    </div>
                    <div class="sn-pivot-settings-panel" data-panel="layout">
                        <div class="sn-pivot-settings-section">
                            <div class="sn-pivot-settings-section-title">Layout</div>
                            <label class="sn-pivot-settings-option">
                                <input type="radio" name="pivot-layout" value="compact"${layoutMode === 'compact' ? ' checked' : ''}>
                                <span>Compact Form</span>
                            </label>
                            <label class="sn-pivot-settings-option">
                                <input type="radio" name="pivot-layout" value="outline"${layoutMode === 'outline' ? ' checked' : ''}>
                                <span>Outline Form</span>
                            </label>
                            <label class="sn-pivot-settings-option">
                                <input type="radio" name="pivot-layout" value="tabular"${layoutMode === 'tabular' ? ' checked' : ''}>
                                <span>Tabular Form</span>
                            </label>
                        </div>
                        <div class="sn-pivot-settings-section">
                            <div class="sn-pivot-settings-section-title">Print</div>
                            <label class="sn-pivot-settings-option">
                                <input type="checkbox" data-field="pivot-repeat-labels"${repeatLabels ? ' checked' : ''}>
                                <span>Repeat All Item Labels</span>
                            </label>
                            <label class="sn-pivot-settings-option">
                                <input type="checkbox" data-field="pivot-insert-blank"${insertBlank ? ' checked' : ''}>
                                <span>Insert Blank Line After Each Item</span>
                            </label>
                        </div>
                    </div>
                </div>
            `,
            confirmKey: 'core.layout.layout.modal.m001.confirm', confirmText: 'OK',
            cancelKey: 'core.layout.layout.modal.m001.cancel', cancelText: 'Cancel',
            onOpen: (bodyEl) => {
                const tabs = bodyEl.querySelectorAll('.sn-pivot-settings-tab');
                const panels = bodyEl.querySelectorAll('.sn-pivot-settings-panel');
                const subtotalSelect = bodyEl.querySelector('[data-field="pivot-subtotal-position"]');
                const subtotalRadios = bodyEl.querySelectorAll('input[name="pivot-subtotal"]');

                const syncSubtotalState = () => {
                    const mode = bodyEl.querySelector('input[name="pivot-subtotal"]:checked')?.value || 'auto';
                    if (subtotalSelect) subtotalSelect.disabled = mode === 'none';
                };

                subtotalRadios.forEach(radio => {
                    radio.onchange = syncSubtotalState;
                });
                syncSubtotalState();

                tabs.forEach(tab => {
                    tab.onclick = () => {
                        const target = tab.dataset.tab;
                        tabs.forEach(t => t.classList.toggle('active', t === tab));
                        panels.forEach(panel => {
                            panel.classList.toggle('active', panel.dataset.panel === target);
                        });
                    };
                });
            },
            onConfirm: (bodyEl) => {
                const nameInput = bodyEl.querySelector('[data-field="pivot-custom-name"]');
                const subtotalModeValue = bodyEl.querySelector('input[name="pivot-subtotal"]:checked')?.value || 'auto';
                const subtotalPosValue = bodyEl.querySelector('[data-field="pivot-subtotal-position"]')?.value || 'top';
                const sortTypeValue = bodyEl.querySelector('[data-field="pivot-sort-type"]')?.value || 'manual';
                const showDropDownsValue = bodyEl.querySelector('[data-field="pivot-show-dropdowns"]')?.checked ?? true;
                const showAllValue = bodyEl.querySelector('[data-field="pivot-show-all"]')?.checked ?? false;
                const repeatLabelsValue = bodyEl.querySelector('[data-field="pivot-repeat-labels"]')?.checked ?? false;
                const insertBlankValue = bodyEl.querySelector('[data-field="pivot-insert-blank"]')?.checked ?? false;
                const layoutValue = bodyEl.querySelector('input[name="pivot-layout"]:checked')?.value || layoutMode;

                const newName = nameInput?.value?.trim() || sourceName;
                field.name = newName;
                field.defaultSubtotal = subtotalModeValue !== 'none';
                field.subtotalTop = subtotalPosValue === 'top';
                field.sortType = sortTypeValue;
                field.showDropDowns = showDropDownsValue;
                field.showAll = showAllValue;
                field.repeatItemLabels = repeatLabelsValue;
                field.insertBlankRow = insertBlankValue;

                switch (layoutValue) {
                    case 'compact':
                        pt.compact = true;
                        pt.compactData = true;
                        pt.outline = true;
                        pt.outlineData = true;
                        break;
                    case 'outline':
                        pt.compact = false;
                        pt.compactData = false;
                        pt.outline = true;
                        pt.outlineData = true;
                        break;
                    case 'tabular':
                        pt.compact = false;
                        pt.compactData = false;
                        pt.outline = false;
                        pt.outlineData = false;
                        break;
                }

                pt._needsRecalculate = true;
                pt.calculate();
                pt.render();
                this._syncPivotPanelState(pt);
            }
        });
    }

    _openPivotValueFieldSettingsModal(pt, dataFieldIndex) {
        const SN = this.SN;
        const df = pt.dataFields[dataFieldIndex];
        if (!df) return;

        const fieldIndex = df.fld;
        const cacheField = pt.cache?.fields?.[fieldIndex];
        const field = pt.pivotFields[fieldIndex];
        const sourceName = cacheField?.name || field?.name || `Field${fieldIndex}`;
        const customName = df.name || df.getDefaultName(sourceName, SN);

        const subtotalOptions = Object.entries(PivotDataField.SUBTOTAL_MAP)
            .map(([key, info]) => `<option value="${key}"${df.subtotal === key ? ' selected' : ''}>${SN.t(info.labelKey)}</option>`)
            .join('');

        const showAsOptions = Object.entries(PivotDataField.SHOW_DATA_AS_MAP)
            .map(([key, info]) => `<option value="${key}"${df.showDataAs === key ? ' selected' : ''}>${SN.t(info.labelKey)}</option>`)
            .join('');

        const baseFieldCandidates = Array.from(new Set([
            ...pt.rowFields,
            ...pt.colFields,
            ...(df.baseField >= 0 ? [df.baseField] : [])
        ]));
        const baseFieldOptions = [
            `<option value="-1">(None)</option>`,
            ...baseFieldCandidates.map(idx => {
                const name = pt.pivotFields[idx]?.name || pt.cache?.fields?.[idx]?.name || `Field${idx}`;
                const selected = df.baseField === idx ? ' selected' : '';
                return `<option value="${idx}"${selected}>${name}</option>`;
            })
        ].join('');

        const getBaseItems = (baseFieldIndex) => {
            if (baseFieldIndex === null || baseFieldIndex === undefined || baseFieldIndex < 0) return [];
            const baseField = pt.pivotFields[baseFieldIndex];
            if (!baseField) return [];
            return baseField.items.map((item, itemIndex) => ({
                index: itemIndex,
                text: pt._getItemDisplayValue(baseFieldIndex, item.x) || ''
            }));
        };

        SN.Utils.modal({
            titleKey: 'core.layout.layout.modal.m002.title', title: 'Value Field Settings',
            maxWidth: '480px',
            content: `
                <div class="sn-pivot-settings">
                    <div class="sn-pivot-settings-row">
                        <label>Source Name</label>
                        <div class="sn-pivot-settings-static">${sourceName}</div>
                    </div>
                    <div class="sn-pivot-settings-row">
                        <label>Custom Name</label>
                        <input type="text" data-field="pivot-value-name" value="${customName}">
                    </div>
                    <div class="sn-pivot-settings-tabs">
                        <button type="button" class="sn-pivot-settings-tab active" data-tab="summary">Summarize Values By</button>
                        <button type="button" class="sn-pivot-settings-tab" data-tab="showas">Show Values As</button>
                    </div>
                    <div class="sn-pivot-settings-panel active" data-panel="summary">
                        <div class="sn-pivot-settings-section">
                            <div class="sn-pivot-settings-section-title">Summarize Values By</div>
                            <select data-field="pivot-value-subtotal">${subtotalOptions}</select>
                        </div>
                    </div>
                    <div class="sn-pivot-settings-panel" data-panel="showas">
                        <div class="sn-pivot-settings-section">
                            <div class="sn-pivot-settings-section-title">Show Values As</div>
                            <div class="sn-pivot-settings-row">
                                <label>Show Values As</label>
                                <select data-field="pivot-value-showas">${showAsOptions}</select>
                            </div>
                            <div class="sn-pivot-settings-row">
                                <label>Base Field</label>
                                <select data-field="pivot-value-base-field">${baseFieldOptions}</select>
                            </div>
                            <div class="sn-pivot-settings-row">
                                <label>Base Item</label>
                                <select data-field="pivot-value-base-item"></select>
                            </div>
                        </div>
                    </div>
                </div>
            `,
            confirmKey: 'core.layout.layout.modal.m002.confirm', confirmText: 'OK',
            cancelKey: 'core.layout.layout.modal.m002.cancel', cancelText: 'Cancel',
            onOpen: (bodyEl) => {
                const tabs = bodyEl.querySelectorAll('.sn-pivot-settings-tab');
                const panels = bodyEl.querySelectorAll('.sn-pivot-settings-panel');
                const showAsSelect = bodyEl.querySelector('[data-field="pivot-value-showas"]');
                const baseFieldSelect = bodyEl.querySelector('[data-field="pivot-value-base-field"]');
                const baseItemSelect = bodyEl.querySelector('[data-field="pivot-value-base-item"]');

                const requiresBaseField = (value) => {
                    return !['normal', 'percentOfTotal', 'percentOfRow', 'percentOfCol'].includes(value);
                };

                const renderBaseItems = () => {
                    const baseFieldIndex = Number(baseFieldSelect.value);
                    const items = getBaseItems(baseFieldIndex);
                    const dynamicOptions = items.map(item => {
                        const selected = df.baseItem === item.index ? ' selected' : '';
                        return `<option value="${item.index}"${selected}>${item.text || '(Blank)'}</option>`;
                    }).join('');
                    const relativeOptions = [
                        `<option value="-1"${df.baseItem === -1 ? ' selected' : ''}>(Previous item)</option>`,
                        `<option value="-2"${df.baseItem === -2 ? ' selected' : ''}>(Next item)</option>`
                    ].join('');
                    baseItemSelect.innerHTML = relativeOptions + dynamicOptions;
                    if (!items.length) {
                        const noneSelected = df.baseItem >= 0 ? ' selected' : '';
                        baseItemSelect.innerHTML += `<option value="0"${noneSelected}>(None)</option>`;
                    }
                };

                const syncBaseControls = () => {
                    const showAsValue = showAsSelect.value;
                    const enableBase = requiresBaseField(showAsValue);
                    baseFieldSelect.disabled = !enableBase;
                    baseItemSelect.disabled = !enableBase;
                    if (enableBase) {
                        if (Number(baseFieldSelect.value) < 0 && baseFieldCandidates.length > 0) {
                            baseFieldSelect.value = String(baseFieldCandidates[0]);
                        }
                        renderBaseItems();
                    } else {
                        baseFieldSelect.value = '-1';
                        baseItemSelect.innerHTML = '<option value="0">(None)</option>';
                    }
                };

                tabs.forEach(tab => {
                    tab.onclick = () => {
                        const target = tab.dataset.tab;
                        tabs.forEach(t => t.classList.toggle('active', t === tab));
                        panels.forEach(panel => {
                            panel.classList.toggle('active', panel.dataset.panel === target);
                        });
                    };
                });

                showAsSelect.onchange = syncBaseControls;
                baseFieldSelect.onchange = renderBaseItems;
                syncBaseControls();
            },
            onConfirm: (bodyEl) => {
                const nameInput = bodyEl.querySelector('[data-field="pivot-value-name"]');
                const subtotalValue = bodyEl.querySelector('[data-field="pivot-value-subtotal"]')?.value || df.subtotal;
                const showAsValue = bodyEl.querySelector('[data-field="pivot-value-showas"]')?.value || df.showDataAs;
                const baseFieldValue = Number(bodyEl.querySelector('[data-field="pivot-value-base-field"]')?.value ?? -1);
                const baseItemValue = Number(bodyEl.querySelector('[data-field="pivot-value-base-item"]')?.value ?? 0);

                df.subtotal = subtotalValue;
                df.showDataAs = showAsValue;

                if (['normal', 'percentOfTotal', 'percentOfRow', 'percentOfCol'].includes(showAsValue)) {
                    df.baseField = -1;
                    df.baseItem = 0;
                } else {
                    df.baseField = baseFieldValue;
                    df.baseItem = baseItemValue;
                }

                const newName = nameInput?.value?.trim();
                df.name = newName ? newName : df.getDefaultName(sourceName, SN);

                pt._needsRecalculate = true;
                pt.calculate();
                pt.render();
                this._syncPivotPanelState(pt);
            }
        });
    }

    _addFieldToZone(pt, zone, index, subtotal = 'sum', position = -1) {
        const field = pt.cache.fields[index];
        if (!field) return;

        switch (zone) {
            case 'row':
                if (!pt.rowFields.includes(index)) {
                    pt.addRowField(index, position);
                }
                break;
            case 'column':
                if (!pt.colFields.includes(index)) {
                    pt.addColField(index, position);
                }
                break;
            case 'value':
                pt.addDataField(index, subtotal || 'sum', null, position);
                break;
            case 'filter':
                if (!pt.pageFields.includes(index)) {
                    if (position === -1 || position >= pt.pageFields.length) {
                        pt.addPageField(index);
                    } else {
                        pt.addPageField(index);
                        this._moveFieldWithinZone(pt, 'filter', index, position);
                    }
                }
                break;
        }
    }

    _removeFieldFromZone(pt, zone, index, dataFieldIndex = null) {
        switch (zone) {
            case 'row':
                pt.removeRowField(index);
                break;
            case 'column':
                pt.removeColField(index);
                break;
            case 'value':
                if (typeof dataFieldIndex === 'number' && !Number.isNaN(dataFieldIndex)) {
                    pt.removeDataField(dataFieldIndex);
                } else {
                    const dfIdx = pt.dataFields.findIndex(df => df.fld === index);
                    if (dfIdx !== -1) {
                        pt.removeDataField(dfIdx);
                    }
                }
                break;
            case 'filter':
                pt.removePageField(index);
                break;
        }
    }

    _moveFieldWithinZone(pt, zone, index, position = -1) {
        let list = null;
        switch (zone) {
            case 'row':
                list = pt.rowFields;
                break;
            case 'column':
                list = pt.colFields;
                break;
            case 'filter':
                list = pt.pageFields;
                break;
            default:
                return;
        }

        const currentIndex = list.indexOf(index);
        if (currentIndex === -1) return;

        let targetIndex = position;
        if (targetIndex < 0 || targetIndex > list.length) targetIndex = list.length;
        if (targetIndex > currentIndex) targetIndex -= 1;

        if (targetIndex === currentIndex) return;
        list.splice(currentIndex, 1);
        list.splice(targetIndex, 0, index);
        pt._needsRecalculate = true;
    }

    _getDefaultFieldZone(pt, index) {
        const cacheField = pt.cache?.fields?.[index];
        const sharedItems = cacheField?.sharedItems;
        if (sharedItems) {
            const hasNumber = sharedItems.containsNumber === true;
            const hasString = sharedItems.containsString === true;
            const hasDate = sharedItems.containsDate === true;
            if (hasNumber && !hasString && !hasDate) return 'value';
        }
        return 'row';
    }

    _getPivotZoneDropIndex(zoneList, event) {
        const items = Array.from(zoneList.querySelectorAll('.sn-pivot-zone-item'));
        if (items.length === 0) return 0;

        const targetItem = event.target.closest('.sn-pivot-zone-item');
        if (!targetItem) return items.length;

        const rect = targetItem.getBoundingClientRect();
        const targetIndex = items.indexOf(targetItem);
        const before = event.clientY < rect.top + rect.height / 2;
        return before ? targetIndex : targetIndex + 1;
    }

    _removeFieldFromAllZones(pt, index) {
        this._removeFieldFromZone(pt, 'row', index);
        this._removeFieldFromZone(pt, 'column', index);
        while (pt.dataFields.some(df => df.fld === index)) {
            this._removeFieldFromZone(pt, 'value', index);
        }
        this._removeFieldFromZone(pt, 'filter', index);
    }

    _updatePivotTableIfNeeded(pt) {
        if (!this._delayUpdate) {
            pt.calculate();
            pt.render();
        }
    }

    // ============================= 工作表Tab滚动 =============================

    _sheetScrollPos = 0;

    /**
     * 滚动工作表标签
     * @param {number} dir - 滚动方向（1 或 -1）
     * @returns {void}
     */
    scrollSheetTabs(dir) {
        const container = this.SN.containerDom.querySelector('.sn-sheet-list');
        const scroll = this.SN.containerDom.querySelector('.sn-sheet-scroll');
        if (!container || !scroll) return;

        const step = 100;
        const maxScroll = Math.min(0, container.clientWidth - scroll.scrollWidth);

        this._sheetScrollPos += dir * -step;
        this._sheetScrollPos = Math.max(maxScroll, Math.min(0, this._sheetScrollPos));
        scroll.style.transform = `translateX(${this._sheetScrollPos}px)`;
    }

    // ============================= 上下文工具栏 =============================

    _currentContextType = null;

    /**
     * 更新上下文工具栏（根据选择内容显示/隐藏上下文选项卡）
     * @param {Object} context - 上下文信息 { type: 'table'|'pivotTable'|'chart'|null, data: any, autoSwitch: boolean }
     *   - autoSwitch: 是否自动切换到上下文选项卡（默认false，只显示不切换）
     * @returns {void}
     */
    updateContextualToolbar(context) {
        const contextType = context?.type || null;
        const autoSwitch = context?.autoSwitch || false;
        this._currentContextData = context?.data || null;

        const menuList = this.SN.containerDom.querySelector('.sn-menu-list');
        if (!menuList) return;

        // 如果上下文类型没有变化且不需要自动切换，只更新信息
        if (this._currentContextType === contextType && !autoSwitch) {
            if (contextType === 'table' && context?.data) {
                this._updateTableDesignPanel(context.data);
            } else if (contextType === 'pivotTable' && context?.data) {
                this._updatePivotDesignPanel(context.data);
            }
            this._scheduleToolbarAdaptiveLayout();
            return;
        }
        this._currentContextType = contextType;

        // 隐藏所有上下文标签
        menuList.querySelectorAll('.sn-contextual-tab').forEach(tab => {
            tab.style.display = 'none';
        });

        if (contextType) {
            // 显示对应的上下文标签（可能有多个，如透视表有分析和设计两个标签）
            const contextTabs = menuList.querySelectorAll(`.sn-contextual-tab[data-context-type="${contextType}"]`);
            contextTabs.forEach((contextTab, index) => {
                contextTab.style.display = '';
                // 只有 autoSwitch=true 时才自动切换到第一个上下文选项卡
                if (autoSwitch && index === 0) {
                    this._switchToTab(contextTab);
                }
            });
            // 更新面板信息
            if (contextType === 'table' && context.data) {
                this._updateTableDesignPanel(context.data);
            } else if (contextType === 'pivotTable' && context.data) {
                this._updatePivotDesignPanel(context.data);
            }
        } else {
            // 移出上下文区域：如果当前激活的是上下文标签，切换回"开始"
            const activeTab = menuList.querySelector('.sn-menu-tab.active');
            if (activeTab?.classList.contains('sn-contextual-tab')) {
                activeTab.classList.remove('active');
                const startTab = menuList.querySelector('.sn-menu-tab[data-panel="start"]');
                if (startTab) this._switchToTab(startTab);
            }
            this.refreshToolbar();
        }
        this._scheduleToolbarAdaptiveLayout();
    }

    _switchToTab(tab) {
        const menuList = tab.parentElement;
        const tools = this.SN.containerDom.querySelector('.sn-tools');
        const panelName = tab.dataset.panel;

        // 更新标签状态
        menuList.querySelectorAll('.sn-menu-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        // 更新面板状态
        tools.querySelectorAll('.sn-tools-panel').forEach(p => p.classList.remove('active'));
        const targetPanel = tools.querySelector(`.sn-tools-panel[data-panel="${panelName}"]`);
        if (targetPanel) {
            // 懒加载面板
            if (this._toolbarBuilder && this._toolbarBuilder.renderPanelIfNeeded(targetPanel)) {
                this.initColorComponents(targetPanel);
            }
            targetPanel.classList.add('active');
            if (this._currentContextType === 'table' && this._currentContextData) {
                this._updateTableDesignPanel(this._currentContextData);
            } else if (this._currentContextType === 'pivotTable' && this._currentContextData) {
                this._updatePivotDesignPanel(this._currentContextData);
            }
        }
        this.refreshToolbar();
        this._scheduleToolbarAdaptiveLayout();
    }

    _updateTableDesignPanel(table) {
        const tools = this.SN.containerDom.querySelector('.sn-tools');
        if (!tools) return;

        // 更新表名称输入框
        const nameInput = tools.querySelector('[data-field="tableName"]');
        if (nameInput) {
            nameInput.value = table.displayName || table.name || '';
        }

        // 更新表选项复选框
        const checkboxMap = {
            'tableHeaderRow': table.showHeaderRow === true,
            'tableTotalsRow': table.showTotalsRow === true,
            'tableRowStripes': table.showRowStripes,
            'tableColStripes': table.showColumnStripes,
            'tableFirstColumn': table.showFirstColumn,
            'tableLastColumn': table.showLastColumn,
            'tableAutoFilter': table.autoFilterEnabled === true
        };

        for (const [id, checked] of Object.entries(checkboxMap)) {
            const checkbox = tools.querySelector(`[data-field="${id}"]`);
            if (checkbox) checkbox.checked = checked;
        }

    }

    _updatePivotAnalyzePanel(pivotTable) {
        if (!pivotTable) return;
        const tools = this.SN.containerDom.querySelector('.sn-tools');
        if (!tools) return;

        const fieldListBtn = tools.querySelector('[data-field="pivotFieldList"]');
        if (fieldListBtn) {
            const isActive = this._showPivotPanel && this._activePivotTable === pivotTable;
            fieldListBtn.classList.toggle('sn-btn-selected', isActive);
        }

        const drillBtn = tools.querySelector('[data-field="pivotDrillButtons"]');
        if (drillBtn) {
            drillBtn.classList.toggle('sn-btn-selected', pivotTable.showDrill !== false);
        }

        const headersBtn = tools.querySelector('[data-field="pivotFieldHeaders"]');
        if (headersBtn) {
            const headersActive = pivotTable.showRowHeaders !== false && pivotTable.showColHeaders !== false;
            headersBtn.classList.toggle('sn-btn-selected', headersActive);
        }
    }

    _updatePivotDesignPanel(pivotTable) {
        const tools = this.SN.containerDom.querySelector('.sn-tools');
        if (!tools) return;

        // 更新透视表名称输入框
        const nameInput = tools.querySelector('[data-field="pivotTableName"]');
        if (nameInput) {
            nameInput.value = pivotTable.name || '';
        }

        // 更新透视表选项复选框
        const checkboxMap = {
            'pivotRowHeaders': pivotTable.showRowHeaders !== false,
            'pivotColHeaders': pivotTable.showColHeaders !== false,
            'pivotRowStripes': pivotTable.showRowStripes === true,
            'pivotColStripes': pivotTable.showColStripes === true
        };

        for (const [id, checked] of Object.entries(checkboxMap)) {
            const checkbox = tools.querySelector(`[data-field="${id}"]`);
            if (checkbox) checkbox.checked = checked;
        }

        this._updatePivotAnalyzePanel(pivotTable);
    }

    // ============================= 私有方法 =============================

    _init(dom) {
        if (this._isInitialized) return;

        // 设置viewport meta标签，禁止缩放，100%显示
        setupViewport();

        // 创建DOM结构
        this._createDOMStructure(dom);

        // 缓存DOM元素引用
        this._cacheDOMElements();

        // 初始化组件和事件
        this._initializeComponents();

        // 初始化下拉
        initDropdown(this.SN.containerDom);

        // 初始化tooltip
        initTooltip(this.SN.containerDom);

        this._isInitialized = true;
    }

    _createDOMStructure(dom) {
        const ns = this.SN.namespace;
        const menuRightHTML = this.menuConfig.right.html;

        // 使用DOMBuilder模块生成HTML（面板标签固定，不可配置）
        const menuListCallback = typeof this._options.menuList === 'function' ? this._options.menuList : null;
        const { html: mainHTML, toolbarBuilder } = DOMBuilder.createMainHTML(ns, menuRightHTML, this.SN, menuListCallback);
        // 保存 builder 实例供懒加载使用，并设置 SN 引用（用于 checkbox getter）
        this._toolbarBuilder = toolbarBuilder;
        toolbarBuilder.setSN(this.SN);

        dom.insertAdjacentHTML('afterbegin', mainHTML);
    }

    _cacheDOMElements() {
        // 缓存所有需要的DOM元素，避免重复查询
        this.snHead = this.SN.containerDom.querySelector('.sn-menu');
        this.snTools = this.SN.containerDom.querySelector('.sn-tools');
        this.snFormulaBar = this.SN.containerDom.querySelector('.sn-input');
        this.snSheetTabBar = this.SN.containerDom.querySelector('.sn-sheets');
        this.snChat = this.SN.containerDom.querySelector('.sn-chat');
        this.snMain = this.SN.containerDom.querySelector('.sn-main');
        this.snOp = this.SN.containerDom.querySelector('.sn-op');
        this.snCanvas = this.SN.containerDom.querySelector('.sn-canvas');
        this.snPivotPanel = this.SN.containerDom.querySelector('.sn-pivot-panel');
    }

    _initializeComponents() {
        // 按顺序初始化所有组件，避免重复初始化
        this.initResizeObserver();
        this.initDragEvents();
        this.initColorComponents();
        if (this.SN.Action?.initToolbarDefaults) {
            this.SN.Action.initToolbarDefaults();
        }
        this.initEventListeners();
        this.removeLoading();

        // 初始化工具栏事件（标签切换 + 懒加载）
        // 传入颜色初始化回调，用于懒加载面板渲染后初始化其中的颜色组件
        initToolsEvents(this.SN.containerDom, this._toolbarBuilder, (panel) => {
            this.initColorComponents(panel);
        });

        this._bindMenuRightEvents();
        this._bindMinimalToolbarEvents();
        this._syncToolbarModeByWidth(this.SN.containerDom.offsetWidth);
        this._scheduleToolbarAdaptiveLayout();
    }
}

// 使用原型链挂载布局处理模块方法
Object.assign(Layout.prototype, {
    // 尺寸调整处理
    initResizeObserver: ResizeHandler.initResizeObserver,
    throttleResize: ResizeHandler.throttleResize,
    handleAllResize: ResizeHandler.handleAllResize,
    updateCanvasSize: ResizeHandler.updateCanvasSize,

    // 拖拽处理
    initDragEvents: DragHandler.initDragEvents,

    // 事件管理
    initEventListeners: EventManager.initEventListeners,
    initColorComponents: EventManager.initColorComponents,
    removeLoading: EventManager.removeLoading
});
