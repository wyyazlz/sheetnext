/**
 * 工具栏构建器
 * 生成Ribbon风格工具栏HTML，使用自带sn-dropdown组件
 */

import { createToolbarConfig, createMenuConfig } from './ToolbarConfig.js';
import { BorderStyles, BorderPositions, getStyleSvg, getPosSvg } from '../../assets/borderSvgs.js';
import ShapeSvgs from '../../assets/shapes.js';
import { CHART_SHORTCUT_MENUS, getLocalizedChartCategories } from '../../components/ChartConfig.js';
import { getChartSvg } from '../../assets/chartSvgs.js';
import { getSelectionFont } from '../Canvas/RichTextEditor.js';
import { TABLE_STYLE_GROUPS, getStylePreviewColors } from '../Table/TableStyle.js';
import { CELL_STYLE_GROUP_META, CELL_STYLE_PRESETS } from '../../action/Style.js';
import getSvg from '../../assets/mainSvgs.js';
import { getIconCanvas } from '../CF/icons.js';
import { ICON_SETS } from '../CF/rules/IconSetRule.js';

export class ToolbarBuilder {
    constructor(ns, SN = null, menuListCallback = null) {
        this.ns = ns;
        this.SN = SN; // SN 实例，用于调用 checked getter
        this.config = createToolbarConfig(ns);
        if (typeof menuListCallback === 'function') {
            this.config = menuListCallback(this.config);
        }
        this.menuConfig = createMenuConfig(ns);
        this._disabledItemUidSeed = 0;
        this._disabledItemUidMap = new WeakMap();
        this._disabledItemByUid = new Map();
        this._enFallbackMap = {
            '\u5185\u6846': 'Inside Borders',
            '\u5916\u6846': 'Outside Borders',
            '\u5de6\u4e0a-\u53f3\u4e0b': 'Diagonal Down',
            '\u5de6\u4e0b-\u53f3\u4e0a': 'Diagonal Up',
            '\u5de6': 'Left',
            '\u53f3': 'Right',
            '\u4e0a': 'Top',
            '\u4e0b': 'Bottom',
            '\u65e0': 'None',
            '\u6d45\u8272': 'Light',
            '\u4e2d\u7b49\u6df1\u6d45': 'Medium',
            '\u6df1\u8272': 'Dark',
            '\u5173\u95ed\u9ad8\u4eae': 'Turn Off Highlight',
            '\u597d\u3001\u5dee\u548c\u9002\u4e2d': 'Good, Bad, and Neutral',
            '\u6570\u636e\u548c\u6a21\u578b': 'Data and Model',
            '\u4e3b\u9898\u5355\u5143\u683c\u6837\u5f0f': 'Themed Cell Styles',
            '\u6570\u5b57\u683c\u5f0f': 'Number Format',
            '\u900f\u89c6\u8868\u6837\u5f0f': 'PivotTable Styles',
            '\u900f\u89c6\u8868\u6837\u5f0f\u9009\u9879': 'PivotTable Style Options',
            '\u8868\u683c\u6837\u5f0f\u9009\u9879': 'Table Style Options',
            '1900\u5e741\u67081\u65e5': '1900/1/1'
        };
        // 默认选中第二个面板（第一个通常是"文件"）
        this.activePanel = this.config[1]?.key || this.config[0]?.key || 'start';
    }

    /** 设置 SN 实例（延迟绑定，用于 checked getter） */
    setSN(SN) {
        this.SN = SN;
    }

    _tr(text) {
        if (typeof text !== 'string') return text;
        return this.SN.t(text);
    }

    _t(key, _fallback = '', params = {}) {
        const translated = this.SN.t(key, params);
        if (translated === key && _fallback) return _fallback;
        return translated;
    }

    _trHtml(html) {
        if (typeof html !== 'string') return html;
        return html;
    }

    _containsChinese(text) {
        return typeof text === 'string' && /[\u4e00-\u9fff]/.test(text);
    }

    _isEnglishLocale() {
        const locale = this.SN?.locale || '';
        return /^en\b/i.test(locale);
    }

    _toEnglishLabelFromKey(key) {
        if (typeof key !== 'string') return '';
        const text = key
            .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
            .replace(/[_-]+/g, ' ')
            .replace(/([A-Za-z])(\d)/g, '$1 $2')
            .replace(/(\d)([A-Za-z])/g, '$1 $2')
            .trim();
        if (!text) return '';
        return text
            .split(/\s+/)
            .map(word => {
                if (!word) return word;
                if (/^[A-Z0-9]+$/.test(word)) return word;
                return `${word.charAt(0).toUpperCase()}${word.slice(1)}`;
            })
            .join(' ');
    }

    _trEn(text, fallback = '') {
        const translated = this._tr(text);
        if (!this._isEnglishLocale()) return translated;
        if (!this._containsChinese(translated)) return translated;
        return fallback || this._enFallbackMap[text] || this._enFallbackMap[translated] || translated;
    }

    _itemTitle(item) {
        if (!item) return '';
        const fallback = item.text || item.label || '';
        if (item.titleKey) {
            const titleText = this._t(item.titleKey);
            if (titleText !== item.titleKey) return titleText;
        }
        if (item.labelKey) return this._t(item.labelKey, fallback);
        return this._trEn(fallback);
    }

    getToolbarItemByDisabledUid(uid) {
        if (typeof uid === 'undefined' || uid === null) return null;
        return this._disabledItemByUid.get(String(uid)) || null;
    }

    /** 生成面板HTML - 按需渲染，只渲染活动面板 */
    buildPanelsOnly() {
        const html = this.config.map(panel => {
            const isActive = panel.key === this.activePanel;
            const content = isActive ? this.buildPanel(panel) : '';
            return `<div class="sn-tools-panel${isActive ? ' active' : ''}" data-panel="${panel.key}"${isActive ? ' data-rendered="true"' : ''}>${content}</div>`;
        }).join('');
        return this._trHtml(html);
    }

    /** 按需渲染单个面板（首次切换时调用） */
    renderPanelIfNeeded(panelEl) {
        if (panelEl.dataset.rendered) return false;
        const key = panelEl.dataset.panel;
        const panel = this.config.find(p => p.key === key);
        if (!panel) return false;
        panelEl.innerHTML = this.buildPanel(panel);
        panelEl.dataset.rendered = 'true';
        return true; // 返回true表示新渲染了面板
    }

    buildPanel(panel) {
        const html = panel.groups.map((group, idx) => this.buildGroup(group, idx === panel.groups.length - 1)).join('');
        return this._trHtml(html);
    }

    buildGroup(group, isLast) {
        // 将相邻的 row 合并到 rows 容器中
        const result = [];
        let rowBuffer = [];

        const flushRows = () => {
            if (rowBuffer.length > 0) {
                result.push(`<div class="sn-tools-rows">${rowBuffer.join('')}</div>`);
                rowBuffer = [];
            }
        };

        for (const item of group.items) {
            if (item.type === 'row') {
                rowBuffer.push(this.buildItem(item));
            } else {
                flushRows();
                result.push(this.buildItem(item));
            }
        }
        flushRows();

        return `<div class="sn-tools-group${isLast ? ' last' : ''}"><div class="sn-tools-group-content">${result.join('')}</div></div>`;
    }

    buildItem(item) {
        switch (item.type) {
            case 'large': return this.buildLargeItem(item);
            case 'stack': return this.buildStack(item);
            case 'row': return this.buildRow(item);
            case 'dropdown': return this.buildDropdown(item);
            case 'splitDropdown': return this.buildSplitDropdown(item);
            case 'colorPicker': return this.buildColorPicker(item);
            case 'splitColorPicker': return this.buildSplitColorPicker(item);
            case 'checkbox': return this.buildCheckbox(item);
            case 'labelInput': return this.buildLabelInput(item);
            default: return this.buildSmallItem(item);
        }
    }

    _buildItemLabel(item, { baseClass = '', withArrow = false } = {}) {
        if (!item) return '';
        if (!item.text && !item.label && !item.labelKey && !item.minimalLabel && !item.minimalLabelKey) return '';

        const showInMinimal = item.minimalText !== false;
        const className = `${baseClass}${showInMinimal ? ' sn-minimal-text-label' : ''}`.trim();
        const classAttr = className ? ` class="${className}"` : '';
        const arrow = withArrow ? '<i class="sn-arrow"></i>' : '';
        const fallbackLabel = item.minimalLabel || item.text || item.label;
        const labelKey = item.minimalLabelKey || item.labelKey || '';
        const labelText = labelKey ? this._t(labelKey, fallbackLabel) : this._tr(fallbackLabel);
        return `<span${classAttr}>${labelText}${arrow}</span>`;
    }

    _resolveItemDisabled(item) {
        if (typeof item?.disabled === 'function') {
            return this.SN ? !!item.disabled(this.SN) : false;
        }
        return !!item?.disabled;
    }

    _getDisabledItemUid(item) {
        if (!item || typeof item !== 'object') return '';
        if (this._disabledItemUidMap.has(item)) {
            return this._disabledItemUidMap.get(item);
        }
        const uid = String(++this._disabledItemUidSeed);
        this._disabledItemUidMap.set(item, uid);
        this._disabledItemByUid.set(uid, item);
        return uid;
    }

    _buildDisabledMeta(item) {
        if (!item || typeof item.disabled === 'undefined') {
            return {
                isDisabled: false,
                disabledClass: '',
                disabledAttr: '',
                disabledGetterAttr: '',
                disabledUidAttr: '',
                ariaDisabledAttr: ''
            };
        }

        const isDisabled = this._resolveItemDisabled(item);
        const uid = this._getDisabledItemUid(item);
        return {
            isDisabled,
            disabledClass: isDisabled ? ' sn-item-disabled' : '',
            disabledAttr: isDisabled ? 'disabled' : '',
            disabledGetterAttr: typeof item.disabled === 'function' ? 'data-disabled-getter="true"' : '',
            disabledUidAttr: `data-disabled-uid="${uid}"`,
            ariaDisabledAttr: `aria-disabled="${isDisabled ? 'true' : 'false'}"`
        };
    }

    /** 大图标 */
    buildLargeItem(item) {
        const fieldAttr = item.id ? `data-field="${item.id}"` : '';
        const disabledMeta = this._buildDisabledMeta(item);
        // 支持 active getter（切换按钮双向绑定）
        let activeClass = '';
        let activeGetterAttr = '';
        if (typeof item.active === 'function') {
            // 仅在 SN 实例可用时调用 getter
            activeClass = (this.SN && item.active(this.SN)) ? ' sn-btn-selected' : '';
            activeGetterAttr = 'data-active-getter="true"';
        }
        if (item.menu) {
            const menuStyle = item.menuWidth ? `style="width:${item.menuWidth}"` : '';
            return `
                <div class="sn-dropdown sn-tools-item large${activeClass}${disabledMeta.disabledClass}" ${fieldAttr} ${activeGetterAttr} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr}>
                    <div class="sn-dropdown-toggle sn-no-icon" title="${this._itemTitle(item)}">
                        ${getSvg(item.icon)}
                        ${this._buildItemLabel(item, { baseClass: 'sn-tools-item-label', withArrow: true })}
                    </div>
                    <ul class="sn-dropdown-menu" data-lazy-menu="${item.menu}" ${menuStyle}></ul>
                </div>`;
        }
        const action = item.action ? `onclick="${item.action}"` : '';
        const dblAction = item.dblAction ? `ondblclick="${item.dblAction}"` : '';
        return `
            <div class="sn-tools-item large${activeClass}${disabledMeta.disabledClass}" ${fieldAttr} ${action} ${dblAction} ${activeGetterAttr} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}">
                ${getSvg(item.icon)}
                ${this._buildItemLabel(item, { baseClass: 'sn-tools-item-label' })}
            </div>`;
    }

    /** 小图标 */
    buildSmallItem(item) {
        const fieldAttr = item.id ? `data-field="${item.id}"` : '';
        const disabledMeta = this._buildDisabledMeta(item);
        const action = item.action ? `onclick="${item.action}"` : '';
        const dblAction = item.dblAction ? `ondblclick="${item.dblAction}"` : '';
        const title = this._itemTitle(item);
        return `<div class="sn-tools-item small${disabledMeta.disabledClass}" ${fieldAttr} ${action} ${dblAction} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${title}">${getSvg(item.icon)}</div>`;
    }

    /** 堆叠布局 */
    buildStack(item) {
        const items = item.items.map(i => {
            if (i.type === 'dropdown') return this.buildStackDropdown(i);
            if (i.type === 'splitDropdown') return this.buildStackSplitDropdown(i);
            if (i.type === 'checkbox') return this.buildCheckbox(i);
            return this.buildStackItem(i);
        }).join('');
        return `<div class="sn-tools-stack">${items}</div>`;
    }

    buildStackItem(item) {
        const fieldAttr = item.id ? `data-field="${item.id}"` : '';
        const disabledMeta = this._buildDisabledMeta(item);
        const action = item.action ? `onclick="${item.action}"` : '';
        const title = this._itemTitle(item);
        const label = this._buildItemLabel(item);
        let activeClass = '';
        let activeGetterAttr = '';
        if (typeof item.active === 'function') {
            activeClass = (this.SN && item.active(this.SN)) ? ' sn-btn-selected' : '';
            activeGetterAttr = 'data-active-getter="true"';
        }
        return `<div class="sn-tools-stack-item${activeClass}${disabledMeta.disabledClass}" ${fieldAttr} ${action} ${activeGetterAttr} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${title}">${getSvg(item.icon)}${label}</div>`;
    }

    buildStackDropdown(item) {
        const label = this._buildItemLabel(item);
        const menuStyle = item.menuWidth ? `style="width:${item.menuWidth}"` : '';
        const disabledMeta = this._buildDisabledMeta(item);
        return `
            <div class="sn-dropdown sn-tools-stack-item${disabledMeta.disabledClass}" ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}">
                <div class="sn-dropdown-toggle sn-no-icon">${getSvg(item.icon)}${label}<i class="sn-arrow"></i></div>
                <ul class="sn-dropdown-menu" data-lazy-menu="${item.menu}" ${menuStyle}></ul>
            </div>`;
    }

    buildStackSplitDropdown(item) {
        const keepOpen = item.keepOpen ? 'data-keep-open="true"' : '';
        const label = this._buildItemLabel(item);
        const disabledMeta = this._buildDisabledMeta(item);
        return `
            <div class="sn-split-btn sn-stack-split${disabledMeta.disabledClass}" ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}">
                <div class="sn-split-btn-main" onclick="${item.action}">${getSvg(item.icon)}${label}</div>
                <div class="sn-dropdown sn-split-btn-arrow" ${keepOpen}>
                    <div class="sn-dropdown-toggle sn-no-icon"><i class="sn-arrow"></i></div>
                    <ul class="sn-dropdown-menu" data-lazy-menu="${item.menu}"></ul>
                </div>
            </div>`;
    }

    /** 行布局 */
    buildRow(item) {
        const items = item.items.map(i => this.buildRowItem(i)).join('');
        return `<div class="sn-tools-row">${items}</div>`;
    }

    buildRowItem(item) {
        if (item.type === 'dropdown') {
            // 有width时用inline样式（带边框），否则用纯图标样式
            return item.width ? this.buildInlineDropdown(item) : this.buildDropdown(item);
        }
        if (item.type === 'splitDropdown') return this.buildSplitDropdown(item);
        if (item.type === 'colorPicker') return this.buildColorPicker(item);
        if (item.type === 'splitColorPicker') return this.buildSplitColorPicker(item);
        const fieldAttr = item.id ? `data-field="${item.id}"` : '';
        const disabledMeta = this._buildDisabledMeta(item);
        const action = item.action ? `onclick="${item.action}"` : '';
        const title = this._itemTitle(item);
        const label = this._buildItemLabel(item);
        return `<div class="sn-tools-item small${disabledMeta.disabledClass}" ${fieldAttr} ${action} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${title}">${getSvg(item.icon)}${label}</div>`;
    }

    /** 行内下拉（带边框，用于字体名、字号等）- 不懒加载，因为 combobox 搜索需要菜单项 */
    buildInlineDropdown(item) {
        const menuHTML = this.buildMenuHTML(item.menu);
        const width = item.width ? `style="width:${item.width}"` : '';
        const keepOpen = item.keepOpen ? 'data-keep-open="true"' : '';
        const menuStyle = item.menuWidth ? `style="width:${item.menuWidth}"` : '';
        const iconHtml = item.icon ? getSvg(item.icon) : '';
        const label = item.labelKey
            ? this._t(item.labelKey, item.text || item.label || '')
            : this._tr(item.text || item.label || '');
        const placeholder = item.placeholderKey
            ? this._t(item.placeholderKey, item.placeholder || '')
            : this._tr(item.placeholder || '');
        const id = item.id || '';
        const disabledMeta = this._buildDisabledMeta(item);
        return `
            <div class="sn-dropdown sn-combobox${disabledMeta.disabledClass}" ${keepOpen} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} data-menu="${item.menu}" title="${this._itemTitle(item)}">
                <div class="sn-combobox-main" ${width}>
                    ${iconHtml}<input type="text" class="sn-combobox-input" value="${label}" data-id="${id}" placeholder="${placeholder}" ${disabledMeta.disabledAttr}>
                </div>
                <div class="sn-dropdown-toggle sn-combobox-arrow sn-no-icon"><i class="sn-arrow"></i></div>
                <ul class="sn-dropdown-menu" ${menuStyle}>${menuHTML}</ul>
            </div>`;
    }

    /** 分裂按钮（左边执行，右边下拉） */
    buildSplitDropdown(item) {
        const keepOpen = item.keepOpen ? 'data-keep-open="true"' : '';
        const menuStyle = item.menuWidth ? `style="width:${item.menuWidth}"` : '';
        const iconHtml = item.icon ? getSvg(item.icon) : '';
        const label = this._buildItemLabel(item);
        const toolAttr = item.icon === 'jurassic_borderAll' ? 'data-tool="border"' : '';
        const disabledMeta = this._buildDisabledMeta(item);
        // border 菜单含颜色控件，不支持懒加载
        const noLazy = item.menu === 'border';
        const menuContent = noLazy ? this.buildMenuHTML(item.menu) : '';
        const lazyAttr = noLazy ? '' : `data-lazy-menu="${item.menu}"`;
        return `
            <div class="sn-split-btn${disabledMeta.disabledClass}" ${toolAttr} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}">
                <div class="sn-split-btn-main" onclick="${item.action}">${iconHtml}${label}</div>
                <div class="sn-dropdown sn-split-btn-arrow" ${keepOpen}>
                    <div class="sn-dropdown-toggle sn-no-icon"><i class="sn-arrow"></i></div>
                    <ul class="sn-dropdown-menu" ${lazyAttr} ${menuStyle}>${menuContent}</ul>
                </div>
            </div>`;
    }

    /** 下拉按钮（纯图标，用于下划线、边框等） */
    buildDropdown(item) {
        const keepOpen = item.keepOpen ? 'data-keep-open="true"' : '';
        const label = this._buildItemLabel(item);
        const disabledMeta = this._buildDisabledMeta(item);
        return `
            <div class="sn-dropdown sn-tools-item small${disabledMeta.disabledClass}" ${keepOpen} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}">
                <div class="sn-dropdown-toggle sn-no-icon">${getSvg(item.icon)}${label}<i class="sn-arrow"></i></div>
                <ul class="sn-dropdown-menu" data-lazy-menu="${item.menu}"></ul>
            </div>`;
    }

    /** 颜色选择器 */
    buildColorPicker(item) {
        const colorType = item.colorType || 'font';
        const actionFn = colorType === 'font' ? 'fontColorBtn' : 'fillColorBtn';
        const disabledMeta = this._buildDisabledMeta(item);
        return `
            <div class="sn-dropdown sn-tools-item small${disabledMeta.disabledClass}" data-keep-open="true" ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}">
                <div class="sn-dropdown-toggle sn-no-icon">${getSvg(item.icon)}<i class="sn-arrow"></i></div>
                <ul class="sn-dropdown-menu" style="width:220px;"><div class="sn-color-select" onclick="${this.ns}.Action.${actionFn}(event)"></div></ul>
            </div>`;
    }

    /** 分裂颜色选择器（左边直接应用，右边打开面板） */
    buildSplitColorPicker(item) {
        const colorType = item.colorType || 'font';
        const actionFn = colorType === 'font' ? 'fontColorBtn' : 'fillColorBtn';
        const applyFn = colorType === 'font' ? 'applyFontColor' : 'applyFillColor';
        const keepOpen = item.keepOpen ? 'data-keep-open="true"' : '';
        const disabledMeta = this._buildDisabledMeta(item);
        return `
            <div class="sn-split-btn${disabledMeta.disabledClass}" ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} data-color-type="${colorType}" title="${this._itemTitle(item)}">
                <div class="sn-split-btn-main" onclick="${this.ns}.Action.${applyFn}()"><span class="sn-color-icon">${getSvg(item.icon)}</span><span class="sn-color-bar"></span></div>
                <div class="sn-dropdown sn-split-btn-arrow" ${keepOpen}>
                    <div class="sn-dropdown-toggle sn-no-icon"><i class="sn-arrow"></i></div>
                    <ul class="sn-dropdown-menu" style="width:220px;"><div class="sn-color-select" onclick="${this.ns}.Action.${actionFn}(event)"></div></ul>
                </div>
            </div>`;
    }

    /** 复选框 - 支持 checked 为 getter 函数 */
    buildCheckbox(item) {
        // checked can be a static boolean or a getter function
        let isChecked = false;
        if (typeof item.checked === 'function') {
            // evaluate getter to read live status
            isChecked = this.SN ? !!item.checked(this.SN) : false;
        } else {
            isChecked = !!item.checked;
        }
        const checked = isChecked ? 'checked' : '';
        const disabledMeta = this._buildDisabledMeta(item);

        const action = item.action ? `onchange="${item.action}"` : '';
        // mark dynamic checkbox bindings for later refresh
        const getterAttr = typeof item.checked === 'function' ? 'data-getter="true"' : '';
        const label = item.labelKey ? this._t(item.labelKey, item.text || item.label || '') : this._tr(item.text || item.label || '');
        return `<label class="sn-tools-checkbox${disabledMeta.disabledClass}" ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}"><input type="checkbox" data-field="${item.id}" ${checked} ${action} ${getterAttr} ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.disabledAttr}><span>${label}</span></label>`;
    }

    buildLabelInput(item) {
        const action = item.action ? `onchange="${item.action}"` : '';
        const disabledMeta = this._buildDisabledMeta(item);
        return `
            <div class="sn-tools-label-input${disabledMeta.disabledClass}" ${disabledMeta.disabledGetterAttr} ${disabledMeta.disabledUidAttr} ${disabledMeta.ariaDisabledAttr} title="${this._itemTitle(item)}">
                <label>${item.labelKey ? this._t(item.labelKey, item.text || item.label || '') : this._tr(item.text || item.label || '')}</label>
                <input type="text" data-field="${item.id}" ${action} ${disabledMeta.disabledAttr}>
            </div>`;
    }

    /** 生成菜单项HTML（使用sn-dropdown格式） */
    buildMenuHTML(menuKey) {
        if (menuKey === 'border') return this.buildBorderMenuHTML();
        if (menuKey === 'condFormat') return this.buildCondFormatMenuHTML();
        if (menuKey === 'shapes') return this.buildShapesMenuHTML();
        if (menuKey === 'tableStyles') return this.buildTableStylesMenuHTML();
        if (menuKey === 'formatAsTable') return this.buildFormatAsTableMenuHTML();
        if (menuKey === 'cellStyles') return this.buildCellStylesMenuHTML();
        if (menuKey === 'highlight') return this.buildHighlightMenuHTML();
        if (menuKey === 'rowCol') return this.buildRowColMenuHTML();
        if (CHART_SHORTCUT_MENUS[menuKey]) return this.buildChartMenuHTML(menuKey);

        const menu = this.menuConfig[menuKey];
        if (!menu) return '';

        const html = menu.map(item => {
            if (item.type === 'divider') return '<li class="sn-dropdown-divider"></li>';
            const style = item.style ? `style="${item.style}"` : '';
            // extra 作为兄弟元素，不嵌套在 span 内，这样 dropdown.js 获取第一个 span 就是纯 label
            const extraText = item.extraKey
                ? this._t(item.extraKey, item.extra || '')
                : (item.extra ? this._trEn(item.extra) : '');
            const extra = extraText ? `<span class="sn-float-end sn-color-999">${extraText}</span>` : '';
            const iconHtml = item.icon ? getSvg(item.icon) : '';
            const label = item.labelKey
                ? this._t(item.labelKey, item.text || item.label || '')
                : this._trEn(item.text || item.label);
            return `<li onclick="${item.action}" ${style} data-value="${label}">${iconHtml}<span>${label}</span>${extra}</li>`;
        }).join('');
        return this._trHtml(html);
    }

    /** 行和列菜单（含二级下拉与自定义输入） */
    buildRowColMenuHTML() {
        const ns = this.ns;
        const insertMenu = `
            <ul class="sn-dropdown-submenu sn-rowcol-submenu">
                <li class="sn-rowcol-insert-item" onclick="${ns}.Action.applyRowColInsertByPos(event, 'above')">
                    <span>在上方插入</span>
                    <input type="number" data-field="rowcol-insert-count" min="1" max="10000" value="1"
                        onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">
                    <span>行</span>
                </li>
                <li class="sn-rowcol-insert-item" onclick="${ns}.Action.applyRowColInsertByPos(event, 'below')">
                    <span>在下方插入</span>
                    <input type="number" data-field="rowcol-insert-count" min="1" max="10000" value="1"
                        onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">
                    <span>行</span>
                </li>
                <li class="sn-rowcol-insert-item" onclick="${ns}.Action.applyRowColInsertByPos(event, 'left')">
                    <span>在左侧插入</span>
                    <input type="number" data-field="rowcol-insert-count" min="1" max="10000" value="1"
                        onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">
                    <span>列</span>
                </li>
                <li class="sn-rowcol-insert-item" onclick="${ns}.Action.applyRowColInsertByPos(event, 'right')">
                    <span>在右侧插入</span>
                    <input type="number" data-field="rowcol-insert-count" min="1" max="10000" value="1"
                        onclick="event.stopPropagation()" onmousedown="event.stopPropagation()">
                    <span>列</span>
                </li>
            </ul>
        `;

        return this._trHtml(`
            <li onclick="${ns}.Action.setRowHeightDialog()" data-value="设置行高"><span>设置行高</span></li>
            <li onclick="${ns}.Action.autoFitRowHeight()" data-value="最合适的行高"><span>最合适的行高</span></li>
            <li onclick="${ns}.Action.setDefaultRowHeightDialog()" data-value="默认行高"><span>默认行高</span></li>
            <li class="sn-dropdown-divider"></li>
            <li onclick="${ns}.Action.setColWidthDialog()" data-value="设置列宽"><span>设置列宽</span></li>
            <li onclick="${ns}.Action.autoFitColWidth()" data-value="最合适的列宽"><span>最合适的列宽</span></li>
            <li onclick="${ns}.Action.setDefaultColWidthDialog()" data-value="默认列宽"><span>默认列宽</span></li>
            <li class="sn-dropdown-divider"></li>
            <li class="sn-has-submenu">
                <span>插入行列</span>
                ${insertMenu}
            </li>
            <li class="sn-dropdown-divider"></li>
            <li onclick="${ns}.Action.deleteRowAction()" data-value="删除行"><span>删除行</span></li>
            <li onclick="${ns}.Action.deleteColAction()" data-value="删除列"><span>删除列</span></li>
            <li class="sn-dropdown-divider"></li>
            <li class="sn-has-submenu">
                <span>隐藏和取消隐藏</span>
                <ul class="sn-dropdown-submenu">
                    <li onclick="${ns}.Action.hideRowAction()" data-value="隐藏行"><span>隐藏行</span></li>
                    <li onclick="${ns}.Action.hideColAction()" data-value="隐藏列"><span>隐藏列</span></li>
                    <li onclick="${ns}.Action.unhideRowAction()" data-value="取消隐藏行"><span>取消隐藏行</span></li>
                    <li onclick="${ns}.Action.unhideColAction()" data-value="取消隐藏列"><span>取消隐藏列</span></li>
                    <li class="sn-dropdown-divider"></li>
                    <li onclick="${ns}.Action.unhideAllRowsAction()" data-value="取消隐藏所有行"><span>取消隐藏所有行</span></li>
                    <li onclick="${ns}.Action.unhideAllColsAction()" data-value="取消隐藏所有列"><span>取消隐藏所有列</span></li>
                </ul>
            </li>
        `);
    }

    /** 图表快捷下拉菜单 */
    buildChartMenuHTML(menuKey) {
        const categoryKey = CHART_SHORTCUT_MENUS[menuKey];
        const chartCategories = getLocalizedChartCategories(this.SN);
        const category = chartCategories.find((item) => item.key === categoryKey);
        if (!category) return '';

        const html = category.items.map((item) => {
            const label = item.text || item.label || this._toEnglishLabelFromKey(item.key);
            return `<li class="sn-chart-menu-item" onclick="${this.ns}.Action.insertChart('${item.key}')" data-value="${label}" title="${label}">
                <span class="sn-chart-menu-label">${label}</span>
                <div class="sn-chart-menu-preview">${getChartSvg(item.key)}</div>
            </li>`;
        }).join('');
        return html;
    }

    /** 条件格式菜单 */
    buildCondFormatMenuHTML() {
        const ns = this.ns;
        const cf = `${ns}.Action.condFormat`;

        // 突出显示单元格规则
        const highlightRules = [
            { labelKey: 'action.condFormat.menu.highlightRules.greaterThan', action: `${cf}('cellIs','greaterThan')` },
            { labelKey: 'action.condFormat.menu.highlightRules.lessThan', action: `${cf}('cellIs','lessThan')` },
            { labelKey: 'action.condFormat.menu.highlightRules.between', action: `${cf}('cellIs','between')` },
            { labelKey: 'action.condFormat.menu.highlightRules.equalTo', action: `${cf}('cellIs','equal')` },
            { labelKey: 'action.condFormat.menu.highlightRules.containsText', action: `${cf}('containsText')` },
            { labelKey: 'action.condFormat.menu.highlightRules.dateOccurring', action: `${cf}('timePeriod')` },
            { labelKey: 'action.condFormat.menu.highlightRules.duplicateValues', action: `${cf}('duplicateValues')` },
        ];

        // 项目选取规则
        const topRules = [
            { labelKey: 'action.condFormat.menu.topRules.top10Items', action: `${cf}('top10','top')` },
            { labelKey: 'action.condFormat.menu.topRules.bottom10Items', action: `${cf}('top10','bottom')` },
            { labelKey: 'action.condFormat.menu.topRules.top10Percent', action: `${cf}('top10','topPercent')` },
            { labelKey: 'action.condFormat.menu.topRules.bottom10Percent', action: `${cf}('top10','bottomPercent')` },
            { labelKey: 'action.condFormat.menu.topRules.aboveAverage', action: `${cf}('aboveAverage','above')` },
            { labelKey: 'action.condFormat.menu.topRules.belowAverage', action: `${cf}('aboveAverage','below')` },
        ];

        // 数据条预设 - 渐变填充
        const dataBarGradient = [
            { color: '#638EC6', text: 'Blue' },
            { color: '#63C384', text: 'Green' },
            { color: '#FF555A', text: 'Red' },
            { color: '#FFB628', text: 'Orange' },
            { color: '#8FAADC', text: 'Light Blue' },
            { color: '#C6EFCE', text: 'Light Green' }
        ];

        // 数据条预设 - 纯色填充
        const dataBarSolid = [
            { color: '#638EC6', text: 'Blue' },
            { color: '#63C384', text: 'Green' },
            { color: '#FF555A', text: 'Red' },
            { color: '#FFB628', text: 'Orange' },
            { color: '#8FAADC', text: 'Light Blue' },
            { color: '#C6EFCE', text: 'Light Green' }
        ];

        // 色阶预设
        const colorScales = [
            { colors: ['#F8696B', '#63BE7B'], text: 'Red-Green' },
            { colors: ['#63BE7B', '#F8696B'], text: 'Green-Red' },
            { colors: ['#F8696B', '#FCFCFF'], text: 'Red-White' },
            { colors: ['#FCFCFF', '#F8696B'], text: 'White-Red' },
            { colors: ['#63BE7B', '#FCFCFF'], text: 'Green-White' },
            { colors: ['#FCFCFF', '#63BE7B'], text: 'White-Green' },
            { colors: ['#5A8AC6', '#FCFCFF'], text: 'Blue-White' },
            { colors: ['#FCFCFF', '#5A8AC6'], text: 'White-Blue' },
            { colors: ['#F8696B', '#FFEB84', '#63BE7B'], text: 'Red-Yellow-Green' },
            { colors: ['#63BE7B', '#FFEB84', '#F8696B'], text: 'Green-Yellow-Red' },
            { colors: ['#5A8AC6', '#FCFCFF', '#F8696B'], text: 'Blue-White-Red' },
            { colors: ['#F8696B', '#FCFCFF', '#5A8AC6'], text: 'Red-White-Blue' }
        ];

        // 图标集预设
        const iconSets = {
            Direction: [
                { name: '3Arrows', text: '3 Arrows' },
                { name: '3ArrowsGray', text: '3 Arrows (Gray)' },
                { name: '3Triangles', text: '3 Triangles' },
                { name: '4Arrows', text: '4 Arrows' },
                { name: '4ArrowsGray', text: '4 Arrows (Gray)' },
                { name: '5Arrows', text: '5 Arrows' },
                { name: '5ArrowsGray', text: '5 Arrows (Gray)' }
            ],
            Shapes: [
                { name: '3TrafficLights1', text: 'Traffic Lights' },
                { name: '3TrafficLights2', text: 'Traffic Lights (Rimmed)' },
                { name: '4TrafficLights', text: '4 Traffic Lights' },
                { name: '3Signs', text: '3 Signs' },
                { name: '4RedToBlack', text: 'Red to Black' }
            ],
            Indicators: [
                { name: '3Symbols', text: '3 Symbols' },
                { name: '3Symbols2', text: '3 Circled Symbols' },
                { name: '3Flags', text: '3 Flags' }
            ],
            Rating: [
                { name: '3Stars', text: '3 Stars' },
                { name: '4Rating', text: '4 Ratings' },
                { name: '5Rating', text: '5 Ratings' },
                { name: '5Quarters', text: '5 Pies' },
                { name: '5Boxes', text: '5 Boxes' }
            ]
        };
        const iconSetCategoryLabelKeys = {
            Direction: 'action.condFormat.menu.iconSetCategory.direction',
            Shapes: 'action.condFormat.menu.iconSetCategory.shapes',
            Indicators: 'action.condFormat.menu.iconSetCategory.indicators',
            Rating: 'action.condFormat.menu.iconSetCategory.rating'
        };

        // 渲染函数
        const renderSubmenu = (items) => items.map(i =>
            `<li onclick="${i.action}"><span>${this._t(i.labelKey)}</span></li>`
        ).join('');

        // 渲染数据条预览
        const renderDataBar = (items, gradient) => {
            const type = gradient ? 'gradient' : 'solid';
            return `<div class="sn-cf-databar-grid">${items.map(i =>
                `<div class="sn-cf-databar-item" onclick="${cf}('dataBar','${type}','${i.color}')" title="${this._tr(i.text)}">
                    <div class="sn-cf-databar-preview" style="background:${gradient ? `linear-gradient(to right, ${i.color}, ${i.color}00)` : i.color}"></div>
                </div>`
            ).join('')}</div>`;
        };

        // 渲染色阶预览
        const renderColorScale = (items) => {
            return `<div class="sn-cf-colorscale-grid">${items.map((i, idx) => {
                const gradient = i.colors.length === 2
                    ? `linear-gradient(to right, ${i.colors[0]}, ${i.colors[1]})`
                    : `linear-gradient(to right, ${i.colors[0]}, ${i.colors[1]}, ${i.colors[2]})`;
                return `<div class="sn-cf-colorscale-item" onclick="${cf}('colorScale',${JSON.stringify(i.colors).replace(/"/g, "'")})" title="${this._tr(i.text)}">
                    <div class="sn-cf-colorscale-preview" style="background:${gradient}"></div>
                </div>`;
            }).join('')}</div>`;
        };

        // 渲染图标集预览（使用真实图标）
        const renderIconSet = (category, sets) => {
            const dpr = window.devicePixelRatio || 1;
            const drawSize = Math.round(16 * dpr);
            const categoryKey = iconSetCategoryLabelKeys[category];
            const categoryLabel = categoryKey ? this._t(categoryKey, category) : category;
            return `<div class="sn-cf-iconset-category">
                <div class="sn-cf-iconset-label">${categoryLabel}</div>
                <div class="sn-cf-iconset-grid">${sets.map(s => {
                    const setConfig = ICON_SETS[s.name];
                    if (!setConfig) return '';
                    const icons = [...setConfig.icons].reverse();
                    const imgs = icons.map(n => `<img src="${getIconCanvas(n, drawSize).toDataURL()}">`).join('');
                    return `<div class="sn-cf-iconset-item" onclick="${cf}('iconSet','${s.name}')" title="${this._tr(s.text)}">
                        <div class="sn-cf-iconset-preview">${imgs}</div>
                    </div>`;
                }).join('')}</div>
            </div>`;
        };

        return this._trHtml(`
            <li class="sn-has-submenu"><span>${this._t('action.condFormat.menu.group.highlightCellRules')}</span><i class="sn-arrow-right"></i>
                <ul class="sn-dropdown-submenu">${renderSubmenu(highlightRules)}</ul>
            </li>
            <li class="sn-has-submenu"><span>${this._t('action.condFormat.menu.group.topBottomRules')}</span><i class="sn-arrow-right"></i>
                <ul class="sn-dropdown-submenu">${renderSubmenu(topRules)}</ul>
            </li>
            <li class="sn-dropdown-divider"></li>
            <li class="sn-has-submenu"><span>${this._t('action.condFormat.menu.group.dataBar')}</span><i class="sn-arrow-right"></i>
                <ul class="sn-dropdown-submenu sn-cf-submenu" style="width:200px">
                    <li class="sn-cf-section-title">${this._t('action.condFormat.menu.section.gradientFill')}</li>
                    ${renderDataBar(dataBarGradient, true)}
                    <li class="sn-cf-section-title">${this._t('action.condFormat.menu.section.solidFill')}</li>
                    ${renderDataBar(dataBarSolid, false)}
                    <li class="sn-dropdown-divider"></li>
                    <li onclick="${cf}('dataBar','custom')"><span>${this._t('action.condFormat.menu.item.moreRules')}</span></li>
                </ul>
            </li>
            <li class="sn-has-submenu"><span>${this._t('action.condFormat.menu.group.colorScale')}</span><i class="sn-arrow-right"></i>
                <ul class="sn-dropdown-submenu sn-cf-submenu" style="width:220px">
                    ${renderColorScale(colorScales)}
                    <li class="sn-dropdown-divider"></li>
                    <li onclick="${cf}('colorScale','custom')"><span>${this._t('action.condFormat.menu.item.moreRules')}</span></li>
                </ul>
            </li>
            <li class="sn-has-submenu"><span>${this._t('action.condFormat.menu.group.iconSet')}</span><i class="sn-arrow-right"></i>
                <ul class="sn-dropdown-submenu sn-cf-submenu sn-cf-iconset-menu" style="width:240px">
                    ${Object.entries(iconSets).map(([cat, sets]) => renderIconSet(cat, sets)).join('')}
                    <li class="sn-dropdown-divider"></li>
                    <li onclick="${cf}('iconSet','custom')"><span>${this._t('action.condFormat.menu.item.moreRules')}</span></li>
                </ul>
            </li>
            <li class="sn-dropdown-divider"></li>
            <li onclick="${cf}('newRule')"><span>${this._t('action.condFormat.menu.item.newRule')}</span></li>
            <li class="sn-has-submenu"><span>${this._t('action.condFormat.menu.group.clearRules')}</span><i class="sn-arrow-right"></i>
                <ul class="sn-dropdown-submenu" style="width:180px">
                    <li onclick="${cf}('clearSelection')"><span>${this._t('action.condFormat.menu.item.clearSelection')}</span></li>
                    <li onclick="${cf}('clearSheet')"><span>${this._t('action.condFormat.menu.item.clearSheet')}</span></li>
                </ul>
            </li>
            <li onclick="${cf}('manage')"><span>${this._t('action.condFormat.menu.item.manageRules')}</span></li>
        `);
    }

    /** 边框菜单 */
    buildBorderMenuHTML() {
        const ns = this.ns;
        // 位置图标 (10种)
        const posItems = [
            { key: 'none', action: `${ns}.Action.applyBorderPreset('none')` },
            { key: 'all', action: `${ns}.Action.applyBorderPreset('all')` },
            { key: 'outer', action: `${ns}.Action.applyBorderPreset('outer')` },
            { key: 'inner', action: `${ns}.Action.applyBorderPreset('inner')` },
            { key: 'left', action: `${ns}.Action.applyBorderPreset('left')` },
            { key: 'top', action: `${ns}.Action.applyBorderPreset('top')` },
            { key: 'right', action: `${ns}.Action.applyBorderPreset('right')` },
            { key: 'bottom', action: `${ns}.Action.applyBorderPreset('bottom')` },
            { key: 'diagonalDown', action: `${ns}.Action.applyBorderPreset('diagonalDown')` },
            { key: 'diagonalUp', action: `${ns}.Action.applyBorderPreset('diagonalUp')` },
        ];
        const posHtml = posItems.map(p => {
            const posMeta = BorderPositions[p.key] || {};
            const label = posMeta.labelKey
                ? this._t(posMeta.labelKey)
                : this._toEnglishLabelFromKey(p.key);
            return `<div class="sn-border-pos-item" onclick="${p.action}" title="${label}">${getPosSvg(p.key, 22)}</div>`;
        }).join('');
        // 样式列表 (13种)
        const styleItems = ['hair', 'dotted', 'dashDotDot', 'dashDot', 'slantDashDot', 'dashed', 'thin', 'mediumDashDotDot', 'mediumDashDot', 'mediumDashed', 'medium', 'double', 'thick'];
        const styleHtml = styleItems.map(s =>
            `<li class="sn-border-style-item" onclick="${ns}.Action.applyBorderStyle('${s}')">${getStyleSvg(s, 100, 18)}</li>`
        ).join('');
        return this._trHtml(`
            <div class="sn-border-pos-wrap"><div class="sn-border-pos-grid">${posHtml}</div></div>
            <li class="sn-dropdown-divider"></li>
            <li class="sn-has-submenu"><span>${this._t('action.style.border.menu.color')}</span><i class="sn-arrow-right"></i>
                <ul style="width:220px;" class="sn-dropdown-submenu sn-border-color-submenu"><div class="sn-color-select" onclick="${ns}.Action.borderColorBtn(event)"></div></ul>
            </li>
            <li class="sn-has-submenu"><span>${this._t('action.style.border.menu.style')}</span><i class="sn-arrow-right"></i>
                <ul style="width:120px;min-width:120px" class="sn-dropdown-submenu sn-border-style-submenu">${styleHtml}</ul>
            </li>
        `);
    }

    /** 形状菜单 - 复刻Excel形状下拉列表 */
    buildShapesMenuHTML() {
        const ns = this.ns;

        // SVG显示key到标准shapeType的映射（线条类型需要映射）
        // shapes.js中的SVG key可能与Excel标准shapeType不同，这里做转换
        // Excel标准只有4种连接线：line, straightConnector1, bentConnector3, curvedConnector3
        // 箭头样式通过 arrows 参数控制，而非不同的shapeType
        const connectorMapping = {
            'line': { shapeType: 'line', arrows: 'none' },                           // 直线（无箭头）
            'lineArrowEnd': { shapeType: 'straightConnector1', arrows: 'end' },      // 单箭头直线
            'lineArrowBoth': { shapeType: 'straightConnector1', arrows: 'both' },    // 双箭头直线
            'bentConnector3': { shapeType: 'bentConnector3', arrows: 'none' },       // 肘形连接符（无箭头）
            'bentConnector4': { shapeType: 'bentConnector3', arrows: 'end' },        // 肘形连接符（单箭头）
            'bentConnector5': { shapeType: 'bentConnector3', arrows: 'both' },       // 肘形连接符（双箭头）
            'curvedConnector3': { shapeType: 'curvedConnector3', arrows: 'none' },   // 曲线连接符（无箭头）
            'curvedConnector4': { shapeType: 'curvedConnector3', arrows: 'end' },    // 曲线连接符（单箭头）
            'curvedConnector5': { shapeType: 'curvedConnector3', arrows: 'both' }    // 曲线连接符（双箭头）
        };

        // Excel标准形状分类配置（使用SVG key，渲染时转换为标准shapeType）
        const shapeCategories = [
            {
                name: '线条',
                shapes: ['line', 'lineArrowEnd', 'lineArrowBoth', 'bentConnector3', 'bentConnector4', 'bentConnector5',
                         'curvedConnector3', 'curvedConnector4', 'curvedConnector5']
            },
            {
                name: '矩形',
                shapes: ['rect', 'roundRect', 'snip1Rect', 'snip2SameRect', 'snip2DiagRect',
                         'snipRoundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect']
            },
            {
                name: '基本形状',
                shapes: ['ellipse', 'triangle', 'rtTriangle', 'parallelogram', 'trapezoid', 'diamond',
                         'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon',
                         'pie', 'chord', 'teardrop', 'frame', 'halfFrame', 'corner', 'diagStripe',
                         'plus', 'plaque', 'can', 'cube', 'bevel', 'donut', 'noSmoking', 'blockArc',
                         'foldedCorner', 'smileyFace', 'heart', 'lightningBolt', 'sun', 'moon',
                         'cloud', 'arc', 'bracketPair', 'bracePair', 'leftBracket', 'rightBracket',
                         'leftBrace', 'rightBrace']
            },
            {
                name: '箭头总汇',
                shapes: ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow',
                         'quadArrow', 'leftRightUpArrow', 'bentArrow', 'uturnArrow', 'leftUpArrow', 'bentUpArrow',
                         'curvedRightArrow', 'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow',
                         'stripedRightArrow', 'notchedRightArrow', 'homePlate', 'chevron',
                         'rightArrowCallout', 'downArrowCallout', 'leftArrowCallout', 'upArrowCallout',
                         'leftRightArrowCallout', 'quadArrowCallout', 'circularArrow']
            },
            {
                name: '公式形状',
                shapes: ['mathPlus', 'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual']
            },
            {
                name: '流程图',
                shapes: ['flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision', 'flowChartInputOutput',
                         'flowChartPredefinedProcess', 'flowChartInternalStorage', 'flowChartDocument',
                         'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation',
                         'flowChartManualInput', 'flowChartManualOperation', 'flowChartConnector',
                         'flowChartOffpageConnector', 'flowChartPunchedCard', 'flowChartPunchedTape',
                         'flowChartSummingJunction', 'flowChartOr', 'flowChartCollate', 'flowChartSort',
                         'flowChartExtract', 'flowChartMerge', 'flowChartOnlineStorage', 'flowChartDelay',
                         'flowChartMagneticTape', 'flowChartMagneticDisk', 'flowChartMagneticDrum',
                         'flowChartDisplay']
            },
            {
                name: '星与旗帜',
                shapes: ['irregularSeal1', 'irregularSeal2', 'star4', 'star5', 'star6', 'star7', 'star8',
                         'star10', 'star12', 'star16', 'star24', 'star32',
                         'ribbon2', 'ribbon', 'ellipseRibbon2', 'ellipseRibbon',
                         'verticalScroll', 'horizontalScroll', 'wave', 'doubleWave']
            },
            {
                name: '标注',
                shapes: ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout',
                         'callout1', 'callout2', 'callout3',
                         'borderCallout1', 'borderCallout2', 'borderCallout3',
                         'accentCallout1', 'accentCallout2', 'accentCallout3',
                         'accentBorderCallout1', 'accentBorderCallout2', 'accentBorderCallout3']
            }
        ];

        // 形状名称映射（中文）
        const shapeNames = {
            // 线条（SVG key名称）
            line: '直线', lineArrowEnd: '箭头', lineArrowBoth: '双箭头',
            bentConnector3: '肘形连接符', bentConnector4: '肘形箭头连接符', bentConnector5: '肘形双箭头连接符',
            curvedConnector3: '曲线连接符', curvedConnector4: '曲线箭头连接符', curvedConnector5: '曲线双箭头连接符',
            // 矩形
            rect: '矩形', roundRect: '圆角矩形', round1Rect: '单圆角矩形',
            round2SameRect: '同侧圆角矩形', round2DiagRect: '对角圆角矩形',
            snip1Rect: '剪去单角矩形', snip2SameRect: '剪去同侧角矩形',
            snip2DiagRect: '剪去对角矩形', snipRoundRect: '剪去圆角矩形',
            // 基本形状
            ellipse: '椭圆', triangle: '等腰三角形', rtTriangle: '直角三角形',
            parallelogram: '平行四边形', trapezoid: '梯形', diamond: '菱形',
            pentagon: '五边形', hexagon: '六边形', heptagon: '七边形',
            octagon: '八边形', decagon: '十边形', dodecagon: '十二边形',
            pie: '饼形', chord: '弦形', teardrop: '泪滴形',
            frame: '框架', halfFrame: '半框', corner: 'L形', diagStripe: '斜纹',
            plus: '十字形', plaque: '缺角矩形', can: '圆柱形', cube: '立方体',
            bevel: '棱台', donut: '圆环', noSmoking: '禁止符', blockArc: '空心弧',
            foldedCorner: '折角形', smileyFace: '笑脸', heart: '心形',
            lightningBolt: '闪电形', sun: '太阳形', moon: '月亮形', cloud: '云形',
            arc: '弧形', bracketPair: '双括号', bracePair: '大括号',
            leftBracket: '左中括号', rightBracket: '右中括号',
            leftBrace: '左大括号', rightBrace: '右大括号',
            // 箭头总汇
            rightArrow: '右箭头', leftArrow: '左箭头', upArrow: '上箭头', downArrow: '下箭头',
            leftRightArrow: '左右箭头', upDownArrow: '上下箭头', quadArrow: '四向箭头',
            leftRightUpArrow: '丁字箭头', bentArrow: '直角箭头', uturnArrow: 'U形箭头',
            leftUpArrow: '左上箭头', bentUpArrow: '直角上箭头',
            curvedRightArrow: '右弯箭头', curvedLeftArrow: '左弯箭头',
            curvedUpArrow: '上弯箭头', curvedDownArrow: '下弯箭头',
            stripedRightArrow: '虚尾箭头', notchedRightArrow: '燕尾箭头',
            homePlate: '五边箭头', chevron: 'V形箭头',
            rightArrowCallout: '右箭头标注', leftArrowCallout: '左箭头标注',
            upArrowCallout: '上箭头标注', downArrowCallout: '下箭头标注',
            leftRightArrowCallout: '左右箭头标注', quadArrowCallout: '四向箭头标注',
            circularArrow: '环形箭头',
            // 公式形状
            mathPlus: '加号', mathMinus: '减号', mathMultiply: '乘号',
            mathDivide: '除号', mathEqual: '等号', mathNotEqual: '不等号',
            // 流程图
            flowChartProcess: '过程', flowChartAlternateProcess: '可选过程',
            flowChartDecision: '决策', flowChartInputOutput: '数据',
            flowChartPredefinedProcess: '预定义过程', flowChartInternalStorage: '内部存储',
            flowChartDocument: '文档', flowChartMultidocument: '多文档',
            flowChartTerminator: '终止', flowChartPreparation: '准备',
            flowChartManualInput: '手动输入', flowChartManualOperation: '手动操作',
            flowChartConnector: '接点', flowChartOffpageConnector: '离页连接符',
            flowChartPunchedCard: '卡片', flowChartPunchedTape: '纸带',
            flowChartSummingJunction: '汇总连接', flowChartOr: '或',
            flowChartCollate: '整理', flowChartSort: '排序',
            flowChartExtract: '摘录', flowChartMerge: '合并',
            flowChartOnlineStorage: '存储数据', flowChartDelay: '延期',
            flowChartMagneticTape: '磁带', flowChartMagneticDisk: '磁盘',
            flowChartMagneticDrum: '直接访问存储器', flowChartDisplay: '显示',
            // 星与旗帜
            irregularSeal1: '爆炸形1', irregularSeal2: '爆炸形2',
            star4: '四角星', star5: '五角星', star6: '六角星', star7: '七角星',
            star8: '八角星', star10: '十角星', star12: '十二角星',
            star16: '十六角星', star24: '二十四角星', star32: '三十二角星',
            ribbon: '前凸带形', ribbon2: '上凸带形',
            ellipseRibbon: '上凸弯带形', ellipseRibbon2: '下凸弯带形',
            verticalScroll: '竖卷形', horizontalScroll: '横卷形',
            wave: '波形', doubleWave: '双波形',
            // 标注
            wedgeRectCallout: '矩形标注', wedgeRoundRectCallout: '圆角矩形标注',
            wedgeEllipseCallout: '椭圆形标注', cloudCallout: '云形标注',
            callout1: '线形标注1', callout2: '线形标注2', callout3: '线形标注3',
            borderCallout1: '线形标注1(带边框)', borderCallout2: '线形标注2(带边框)', borderCallout3: '线形标注3(带边框)',
            accentCallout1: '线形标注1(强调线)', accentCallout2: '线形标注2(强调线)', accentCallout3: '线形标注3(强调线)',
            accentBorderCallout1: '线形标注1(带强调边框)', accentBorderCallout2: '线形标注2(带强调边框)', accentBorderCallout3: '线形标注3(带强调边框)'
        };

        // 渲染单个形状项
        const renderShapeItem = (svgKey) => {
            const svg = ShapeSvgs[svgKey];
            if (!svg) return '';
            const name = this._trEn(shapeNames[svgKey] || svgKey, this._toEnglishLabelFromKey(svgKey));

            // 线条类型需要转换为标准shapeType + arrows参数
            const mapping = connectorMapping[svgKey];
            let onclick;
            if (mapping) {
                const arrowsParam = mapping.arrows !== 'none' ? `,{arrows:'${mapping.arrows}'}` : '';
                onclick = `${ns}.activeSheet.Drawing.addShape('${mapping.shapeType}'${arrowsParam});${ns}._r()`;
            } else {
                onclick = `${ns}.activeSheet.Drawing.addShape('${svgKey}');${ns}._r()`;
            }

            return `<div class="sn-shape-item" onclick="${onclick}" title="${name}">
                <svg viewBox="0 0 22 22" width="24" height="24" fill="none" stroke="currentColor" stroke-width="1">${svg}</svg>
            </div>`;
        };

        // 渲染分类
        const categoryFallbacks = ['Lines', 'Rectangles', 'Basic Shapes', 'Block Arrows', 'Equation Shapes', 'Flowchart', 'Stars and Banners', 'Callouts'];
        const renderCategory = (category, idx) => {
            const items = category.shapes
                .filter(s => ShapeSvgs[s])
                .map(renderShapeItem)
                .join('');
            if (!items) return '';
            return `<div class="sn-shape-category">
                <div class="sn-shape-category-title">${this._trEn(category.name, categoryFallbacks[idx] || this._toEnglishLabelFromKey(category.name))}</div>
                <div class="sn-shape-grid">${items}</div>
            </div>`;
        };

        return this._trHtml(`<div class="sn-shapes-menu">${shapeCategories.map((category, idx) => renderCategory(category, idx)).join('')}</div>`);
    }

    /** 表格样式菜单 */
    buildTableStylesMenuHTML() {
        const ns = this.ns;

        // 渲染单个样式预览项
        const renderStyleItem = (styleId) => {
            const colors = getStylePreviewColors(styleId, this.SN);
            if (!colors) return '';

            // 创建 4x3 的迷你表格预览
            return `
                <div class="sn-table-style-item" onclick="${ns}.Action.changeTableStyle('${styleId}')" title="${styleId}">
                    <div class="sn-table-style-preview">
                        <div class="sn-table-style-row header" style="background:${colors.header};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row" style="background:${colors.stripe1};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row stripe" style="background:${colors.stripe2};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row" style="background:${colors.stripe1};border-color:${colors.border}"></div>
                    </div>
                </div>`;
        };

        // 渲染样式分组
        const renderGroup = (title, styles, fallbackTitle = '') => {
            const items = styles.map(renderStyleItem).join('');
            return `
                <div class="sn-table-style-group">
                    <div class="sn-table-style-group-title">${this._trEn(title, fallbackTitle)}</div>
                    <div class="sn-table-style-grid">${items}</div>
                </div>`;
        };

        return this._trHtml(`
            <div class="sn-table-styles-menu sn-scrollbar">
                ${renderGroup('浅色', TABLE_STYLE_GROUPS.light, 'Light')}
                ${renderGroup('中等深浅', TABLE_STYLE_GROUPS.medium, 'Medium')}
                ${renderGroup('深色', TABLE_STYLE_GROUPS.dark, 'Dark')}
                <li class="sn-dropdown-divider"></li>
                <li onclick="${ns}.Action.clearTableStyle()"><span>清除表格样式</span></li>
            </div>`);
    }

    /** 套用表格格式菜单（开始面板） */
    buildFormatAsTableMenuHTML() {
        const ns = this.ns;

        const renderStyleItem = (styleId) => {
            const colors = getStylePreviewColors(styleId, this.SN);
            if (!colors) return '';
            return `
                <div class="sn-table-style-item" onclick="${ns}.Action.formatAsTable('${styleId}')" title="${styleId}">
                    <div class="sn-table-style-preview">
                        <div class="sn-table-style-row header" style="background:${colors.header};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row" style="background:${colors.stripe1};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row stripe" style="background:${colors.stripe2};border-color:${colors.border}"></div>
                        <div class="sn-table-style-row" style="background:${colors.stripe1};border-color:${colors.border}"></div>
                    </div>
                </div>`;
        };

        const renderGroup = (title, styles, fallbackTitle = '') => `
            <div class="sn-table-style-group">
                <div class="sn-table-style-group-title">${this._trEn(title, fallbackTitle)}</div>
                <div class="sn-table-style-grid">${styles.map(renderStyleItem).join('')}</div>
            </div>`;

        return this._trHtml(`
            <div class="sn-table-styles-menu sn-scrollbar">
                ${renderGroup('浅色', TABLE_STYLE_GROUPS.light, 'Light')}
                ${renderGroup('中等深浅', TABLE_STYLE_GROUPS.medium, 'Medium')}
                ${renderGroup('深色', TABLE_STYLE_GROUPS.dark, 'Dark')}
                <li class="sn-dropdown-divider"></li>
                <li onclick="${ns}.Action.clearTableStyle()"><span>清除表格样式</span></li>
            </div>`);
    }

    /** 单元格样式菜单 */
    buildCellStylesMenuHTML() {
        const ns = this.ns;

        const renderStyleItem = (style) => {
            const bg = style.fill?.fgColor || '#FFFFFF';
            const color = style.font?.color || '#000000';
            const fw = style.font?.bold ? 'font-weight:bold;' : '';
            const fs = style.font?.italic ? 'font-style:italic;' : '';
            const sz = style.font?.size ? `font-size:${Math.min(style.font.size, 11)}px;` : '';
            const bd = style.border ? `border:1px solid ${style.border.color || '#ccc'};` : '';
            const label = style.labelKey
                ? this._t(style.labelKey)
                : this._toEnglishLabelFromKey(style.name);

            return `
                <div class="sn-cell-style-item" onclick="${ns}.Action.applyCellStyle('${style.name}')" title="${label}">
                    <div class="sn-cell-style-preview" style="background:${bg};color:${color};${fw}${fs}${sz}${bd}">
                        ${style.numFmt ? '123' : 'Aa'}
                    </div>
                    <div class="sn-cell-style-label">${label}</div>
                </div>`;
        };

        const renderGroup = (groupKey, styles) => {
            if (!styles?.length) return '';
            const groupMeta = CELL_STYLE_GROUP_META[groupKey] || {};
            const groupLabel = groupMeta.labelKey
                ? this._t(groupMeta.labelKey)
                : this._toEnglishLabelFromKey(groupKey);
            return `
                <div class="sn-cell-style-group">
                    <div class="sn-cell-style-group-title">${groupLabel}</div>
                    <div class="sn-cell-style-grid">${styles.map(renderStyleItem).join('')}</div>
                </div>`;
        };

        const groups = Object.entries(CELL_STYLE_PRESETS)
            .map(([groupKey, styles]) => renderGroup(groupKey, styles))
            .join('');

        return this._trHtml(`<div class="sn-cell-styles-menu sn-scrollbar">${groups}</div>`);
    }

    /** 高亮行列颜色面板 */
    buildHighlightMenuHTML() {
        const ns = this.ns;
        const colors = [
            'rgba(255,254,254,0.45)', 'rgba(255,255,204,0.45)', 'rgba(204,255,211,0.45)', 'rgba(230,233,255,0.45)',
            'rgba(255,242,227,0.45)', 'rgba(224,252,255,0.45)', 'rgba(255,242,255,0.45)', 'rgba(221,221,221,0.45)',
            'rgba(255,186,187,0.45)', 'rgba(255,253,186,0.45)', 'rgba(169,255,182,0.45)', 'rgba(214,215,255,0.45)',
            'rgba(255,229,201,0.45)', 'rgba(193,249,255,0.45)', 'rgba(255,213,255,0.45)'
        ];

        const btn = (c) => `<span class="sn-hl-btn" data-color="${c}" style="background:${c.replace(',0.45)', ',1)')}"></span>`;

        return this._trHtml(`
            <div class="sn-highlight-panel" onclick="${ns}.Canvas.setHighlight(event)">
                <div class="sn-hl-row">
                    <span class="sn-hl-btn sn-hl-off" data-color="" title="${this._trEn('Turn Off Highlight', 'Turn Off Highlight')}">✕</span>
                    ${colors.slice(0, 7).map(btn).join('')}
                </div>
                <div class="sn-hl-row">
                    ${colors.slice(7).map(btn).join('')}
                </div>
            </div>`);
    }
}

/** 初始化工具栏事件（标签切换 + 懒加载 + combobox优化 + checkbox状态同步） */
export function initToolsEvents(container, builder, colorInitFn) {
    const tools = container.querySelector('.sn-tools');
    const menuList = container.querySelector('.sn-menu-list');
    if (!tools || !menuList) return;

    const stopDisabledToolEvents = (e) => {
        const disabledHost = e.target.closest('.sn-item-disabled');
        if (!disabledHost || !tools.contains(disabledHost)) return;
        e.preventDefault();
        e.stopImmediatePropagation();
        e.stopPropagation();
    };

    // 在 capture 阶段统一拦截禁用按钮事件，避免 inline onclick/dblclick 执行
    tools.addEventListener('mousedown', stopDisabledToolEvents, true);
    tools.addEventListener('click', stopDisabledToolEvents, true);
    tools.addEventListener('dblclick', stopDisabledToolEvents, true);

    // 刷新面板内 checkbox 状态（面板切换时调用）
    const refreshPanelCheckboxes = (panel) => {
        if (!builder?.SN) return;
        const checkboxes = panel.querySelectorAll('input[type="checkbox"]');
        const config = builder.config;
        checkboxes.forEach(checkbox => {
            const fieldId = checkbox.dataset.field;
            const item = findCheckboxItem(config, fieldId);
            if (!item) return;

            if (typeof item.checked === 'function') {
                checkbox.checked = !!item.checked(builder.SN);
            }

            if (typeof item.disabled === 'function') {
                checkbox.disabled = !!item.disabled(builder.SN);
            } else if (typeof item.disabled !== 'undefined') {
                checkbox.disabled = !!item.disabled;
            }
        });
    };

    // 标签切换 + 面板懒加载
    menuList.addEventListener('click', (e) => {
        const tab = e.target.closest('.sn-menu-tab');
        if (!tab) return;

        const panelName = tab.dataset.panel;

        menuList.querySelectorAll('.sn-menu-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        tools.querySelectorAll('.sn-tools-panel').forEach(p => p.classList.remove('active'));
        const targetPanel = tools.querySelector(`.sn-tools-panel[data-panel="${panelName}"]`);
        if (targetPanel) {
            // 面板懒加载：首次切换时渲染
            if (builder && builder.renderPanelIfNeeded(targetPanel)) {
                // 初始化新面板中的颜色组件
                if (colorInitFn) colorInitFn(targetPanel);
            }
            targetPanel.classList.add('active');
            refreshPanelCheckboxes(targetPanel);
            refreshDisabledToolbarItems(container, builder);
            // 统一刷新工具栏状态（样式、checkbox、active按钮）
            builder?.SN?.Layout?.refreshToolbar?.();
            builder?.SN?.Layout?._scheduleToolbarAdaptiveLayout?.();
        }
    });

    // 懒加载下拉菜单内容（事件委托，在 capture 阶段拦截）
    if (builder) {
        tools.addEventListener('mousedown', (e) => {
            if (e.target.closest('.sn-item-disabled')) return;
            const toggle = e.target.closest('.sn-dropdown-toggle');
            if (!toggle) return;

            const dropdown = toggle.closest('.sn-dropdown');
            if (!dropdown) return;

              const menu = dropdown.querySelector('.sn-dropdown-menu[data-lazy-menu]');
              if (!menu || menu.dataset.lazyLoaded) return;

              // 首次点击，渲染菜单内容
              menu.innerHTML = builder.buildMenuHTML(menu.dataset.lazyMenu);
              menu.dataset.lazyLoaded = 'true';
          }, true);
      }

    // Combobox 增强：打开时标记选中项、选择时设置字体样式
    initComboboxEnhancements(tools, container);
}

/** 初始化 Combobox 增强功能 */
function initComboboxEnhancements(tools, container) {
    // 打开下拉时标记当前选中项
    tools.addEventListener('mousedown', (e) => {
        const combobox = e.target.closest('.sn-combobox');
        if (!combobox) return;
        if (combobox.classList.contains('sn-item-disabled')) return;

        const input = combobox.querySelector('.sn-combobox-input');
        const menu = combobox.querySelector('.sn-dropdown-menu');
        if (!input || !menu) return;

        const currentValue = input.value.trim();
        const customText = container.__sheetNextInstance.t('toolbar.numFmt.custom');

        // 移除所有勾号，添加当前选中项的勾号
        menu.querySelectorAll('li').forEach(item => {
            item.classList.remove('sn-selected');
            const itemValue = item.dataset.value || item.textContent.trim();
            if (itemValue === currentValue) {
                item.classList.add('sn-selected');
            } else if (currentValue === customText && itemValue.startsWith(customText)) {
                item.classList.add('sn-selected');
            }
        });
    }, true);

    // 选择菜单项时：字体设置样式，更新选中状态
    tools.addEventListener('click', (e) => {
        const item = e.target.closest('.sn-combobox .sn-dropdown-menu > li');
        if (!item || item.classList.contains('sn-dropdown-divider')) return;

        const combobox = item.closest('.sn-combobox');
        if (!combobox || combobox.classList.contains('sn-item-disabled')) return;
        const input = combobox?.querySelector('.sn-combobox-input');
        if (!input) return;

        // 字体菜单：设置输入框字体样式
        const menuType = combobox.dataset.menu;
        if (menuType === 'fontName') {
            const value = item.dataset.value || item.querySelector('span')?.textContent?.trim();
            input.style.fontFamily = value;
        }

        // 更新选中状态
        const menu = combobox.querySelector('.sn-dropdown-menu');
        menu.querySelectorAll('li').forEach(li => li.classList.remove('sn-selected'));
        item.classList.add('sn-selected');
    }, true);
}

// ==================== 工具栏状态同步 ====================

// 缓存上次状态（用于差量更新）
let _lastState = null;
let _lastStateStr = '';  // 快速比较用

// 默认值（用于显示，不用于高亮判断）
const DISPLAY_DEFAULTS = {
    fontSize: 11
};

function getDefaultFontName(cell) {
    const locale = cell?._SN?.locale || '';
    return /^zh\b/i.test(locale) ? '宋体' : 'Calibri';
}

// 缓存菜单格式映射（从菜单配置动态生成）
let _numFmtMenuMap = null;

/**
 * 从菜单配置构建格式映射
 * @param {Object} menuConfig - 菜单配置对象
 */
function buildNumFmtMap(menuConfig) {
    if (_numFmtMenuMap) return _numFmtMenuMap;
    _numFmtMenuMap = new Map();

    const numFmtMenu = menuConfig?.numFmt;
    if (!Array.isArray(numFmtMenu)) return _numFmtMenuMap;

    for (const item of numFmtMenu) {
        if (item.type === 'divider' || !item.action) continue;
        // 从 action 中提取格式字符串：areasNumFmt('xxx') 或 areasNumFmt("xxx")
        const match = item.action.match(/areasNumFmt\(['"]([^'"]*)['"]\)/);
        if (match) {
            const fmt = match[1];
            _numFmtMenuMap.set(fmt, item.labelKey || '');
        }
    }
    return _numFmtMenuMap;
}

/**
 * 获取数值格式显示标签
 * 直接从菜单配置匹配，匹配不到显示"自定义"
 */
function getNumFmtLabel(numFmt, menuConfig, t = (key) => key) {
    if (!numFmt || numFmt === 'General') return t('toolbar.numFmt.general');

    const map = buildNumFmtMap(menuConfig);
    const labelKey = map.get(numFmt);
    if (labelKey) return t(labelKey);

    // 尝试标准化后再匹配（处理 ￥ 和 ¥ 的差异）
    const normalized = numFmt.replace(/￥/g, '¥').replace(/&quot;/g, '"');
    const normalizedLabelKey = map.get(normalized);
    if (normalizedLabelKey) return t(normalizedLabelKey);

    return t('toolbar.numFmt.custom');
}

/**
 * 同步工具栏状态（高性能版）
 * - 纯状态差量对比，不依赖 cellKey
 * - 对齐按钮：只有明确设置时才高亮
 * @param {HTMLElement} container - 容器元素
 * @param {Cell} cell - 单元格对象
 * @param {Object} menuConfig - 菜单配置（用于数值格式匹配）
 */
export function syncToolbarState(container, cell, menuConfig) {
    const tools = container.querySelector('.sn-tools');
    if (!tools) return;
    const t = (key) => cell._SN.t(key);

    // 提取当前样式（编辑时从选区获取，否则从单元格获取）
    let font;
    let displayFont;
    if (cell._SN?.Canvas?.inputEditing && !cell.isFormula) {
        font = getSelectionFont(cell._SN.Canvas.input, cell.font) || {};
        displayFont = font;
    } else {
        font = cell.style?.font || {};
        displayFont = cell.font || {};
    }
    const align = cell.style?.alignment || {};
    const fill = cell.fill || {};
    const numFmt = cell.numFmt || 'General';
    const defaultFontColor = cell._SN?.toolbarDefaults?.font;
    const defaultFillColor = cell._SN?.toolbarDefaults?.fill;

    // 构建当前状态
    const currentState = {
        // 显示值（有默认值）
        fontName: font.name || getDefaultFontName(cell),
        fontSize: font.size || DISPLAY_DEFAULTS.fontSize,
        numFmt: getNumFmtLabel(numFmt, menuConfig, t),
        fontColor: displayFont.color || defaultFontColor || '#FF0000',
        fillColor: fill.fgColor || defaultFillColor || '#FFFF00',
        // 开关状态
        bold: !!font.bold,
        italic: !!font.italic,
        underline: font.underline || '',
        strike: !!font.strike,
        // 对齐：空字符串表示默认（不高亮），只有明确设置才高亮
        hAlign: align.horizontal || '',
        vAlign: align.vertical || ''
    };

    // 快速比较：状态字符串相同则跳过
    const currentStateStr = JSON.stringify(currentState);
    if (currentStateStr === _lastStateStr) return;

    // 差量更新
    const updates = [];
    for (const key in currentState) {
        if (!_lastState || currentState[key] !== _lastState[key]) {
            updates.push(key);
        }
    }

    if (updates.length > 0) {
        updateToolbarDOM(tools, currentState, updates);
    }

    _lastState = currentState;
    _lastStateStr = currentStateStr;
}

/** 批量更新工具栏 DOM */
function updateToolbarDOM(tools, state, updates) {
    for (const key of updates) {
        switch (key) {
            case 'fontName': {
                const input = tools.querySelector('.sn-combobox[data-menu="fontName"] .sn-combobox-input');
                if (input) {
                    input.value = state.fontName;
                    input.style.fontFamily = state.fontName;
                }
                break;
            }
            case 'fontSize': {
                const input = tools.querySelector('.sn-combobox[data-menu="fontSize"] .sn-combobox-input');
                if (input) input.value = state.fontSize;
                break;
            }
            case 'bold': {
                const btn = tools.querySelector('.sn-tools-item [data-icon="bold"]')?.closest('.sn-tools-item');
                if (btn) btn.classList.toggle('sn-btn-selected', state.bold);
                break;
            }
            case 'italic': {
                const btn = tools.querySelector('.sn-tools-item [data-icon="italic"]')?.closest('.sn-tools-item');
                if (btn) btn.classList.toggle('sn-btn-selected', state.italic);
                break;
            }
            case 'underline': {
                const btn = tools.querySelector('.sn-tools-item [data-icon="underline"]')?.closest('.sn-tools-item');
                if (btn) btn.classList.toggle('sn-btn-selected', !!state.underline);
                break;
            }
            case 'strike': {
                const btn = tools.querySelector('.sn-tools-item [data-icon="strikethrough"]')?.closest('.sn-tools-item');
                if (btn) btn.classList.toggle('sn-btn-selected', state.strike);
                break;
            }
            case 'fontColor': {
                tools.style.setProperty('--sn-font-color', state.fontColor || 'transparent');
                const btn = tools.querySelector('.sn-split-btn[data-color-type="font"]');
                if (btn) {
                    if (state.fontColor) btn.removeAttribute('data-color-empty');
                    else btn.setAttribute('data-color-empty', 'true');
                }
                break;
            }
            case 'fillColor': {
                tools.style.setProperty('--sn-fill-color', state.fillColor || 'transparent');
                const btn = tools.querySelector('.sn-split-btn[data-color-type="fill"]');
                if (btn) {
                    if (state.fillColor) btn.removeAttribute('data-color-empty');
                    else btn.setAttribute('data-color-empty', 'true');
                }
                break;
            }
            case 'hAlign': {
                // 水平对齐：根据 data-icon 匹配
                const hAlignMap = { left: 'alignLeft', center: 'alignCenter', right: 'alignRight' };
                for (const [align, icon] of Object.entries(hAlignMap)) {
                    const btn = tools.querySelector(`.sn-tools-item [data-icon="${icon}"]`)?.closest('.sn-tools-item');
                    if (btn) btn.classList.toggle('sn-btn-selected', state.hAlign === align);
                }
                break;
            }
            case 'vAlign': {
                // 垂直对齐：根据 data-icon 匹配
                const vAlignMap = { top: 'alignTop', center: 'alignVertically', bottom: 'alignBottom' };
                for (const [align, icon] of Object.entries(vAlignMap)) {
                    const btn = tools.querySelector(`.sn-tools-item [data-icon="${icon}"]`)?.closest('.sn-tools-item');
                    if (btn) btn.classList.toggle('sn-btn-selected', state.vAlign === align);
                }
                break;
            }
            case 'numFmt': {
                const input = tools.querySelector('.sn-combobox[data-menu="numFmt"] .sn-combobox-input');
                if (input) input.value = state.numFmt;
                break;
            }
        }
    }
}

/** 重置缓存（切换工作表时调用） */
export function resetToolbarCache() {
    _lastState = null;
    _lastStateStr = '';
}

/**
 * 刷新工具栏中带有 disabled getter 的项
 * @param {HTMLElement} container - 容器元素
 * @param {ToolbarBuilder} builder - 工具栏构建器实例
 */
export function refreshDisabledToolbarItems(container, builder) {
    if (!container || !builder?.SN) return;
    const tools = container.querySelector('.sn-tools');
    if (!tools) return;

    const targets = tools.querySelectorAll('[data-disabled-getter="true"][data-disabled-uid]');
    targets.forEach(target => {
        const uid = target.dataset.disabledUid;
        if (!uid) return;
        const item = builder.getToolbarItemByDisabledUid(uid);
        if (!item || typeof item.disabled !== 'function') return;

        const isDisabled = !!item.disabled(builder.SN);
        if (target.matches('input[type="checkbox"]')) {
            target.disabled = isDisabled;
            const checkboxLabel = target.closest('.sn-tools-checkbox');
            if (checkboxLabel) {
                checkboxLabel.classList.toggle('sn-item-disabled', isDisabled);
                checkboxLabel.setAttribute('aria-disabled', isDisabled ? 'true' : 'false');
            }
            return;
        }

        if (target.classList.contains('sn-combobox')) {
            const input = target.querySelector('.sn-combobox-input');
            if (input) input.disabled = isDisabled;
        }

        if (target.classList.contains('sn-tools-label-input')) {
            const input = target.querySelector('input, select, textarea');
            if (input) input.disabled = isDisabled;
        }

        target.classList.toggle('sn-item-disabled', isDisabled);
        target.setAttribute('aria-disabled', isDisabled ? 'true' : 'false');

        if (isDisabled) {
            target.classList.remove('active');
            target.querySelectorAll('.sn-dropdown.active').forEach(dropdown => {
                dropdown.classList.remove('active');
            });
        }
    });
}

/**
 * 刷新工具栏中带有 active getter 的按钮状态
 * @param {HTMLElement} container - 容器元素
 * @param {ToolbarBuilder} builder - 工具栏构建器实例
 */
export function refreshActiveButtons(container, builder) {
    if (!container || !builder?.SN) return;
    const tools = container.querySelector('.sn-tools');
    if (!tools) return;

    // 查找所有带有 active getter 的按钮
    const buttons = tools.querySelectorAll('[data-active-getter="true"]');
    buttons.forEach(btn => {
        const fieldId = btn.dataset.field;
        if (!fieldId) return;

        // 从配置中查找对应的配置项
        const item = findActiveItem(builder.config, fieldId);
        if (item && typeof item.active === 'function') {
            const isActive = !!item.active(builder.SN);
            btn.classList.toggle('sn-btn-selected', isActive);
        }
    });
}

export default ToolbarBuilder;

/**
 * 从工具栏配置中查找指定 id 的 checkbox 项
 * @param {Array} config - 工具栏配置
 * @param {string} fieldId - checkbox 的 id
 * @returns {Object|null} checkbox 配置项
 */
function findCheckboxItem(config, fieldId) {
    for (const panel of config) {
        for (const group of panel.groups || []) {
            for (const item of group.items || []) {
                // 直接是 checkbox
                if (item.type === 'checkbox' && item.id === fieldId) {
                    return item;
                }
                // stack 内的 checkbox
                if (item.type === 'stack' && item.items) {
                    const found = item.items.find(i => i.type === 'checkbox' && i.id === fieldId);
                    if (found) return found;
                }
                // row 内的 checkbox
                if (item.type === 'row' && item.items) {
                    const found = item.items.find(i => i.type === 'checkbox' && i.id === fieldId);
                    if (found) return found;
                }
            }
        }
    }
    return null;
}

/**
 * 从工具栏配置中查找指定 id 的带 active getter 的项
 * @param {Array} config - 工具栏配置
 * @param {string} fieldId - 项的 id
 * @returns {Object|null} 配置项
 */
function findActiveItem(config, fieldId) {
    for (const panel of config) {
        for (const group of panel.groups || []) {
            for (const item of group.items || []) {
                // 直接匹配
                if (item.id === fieldId && typeof item.active === 'function') {
                    return item;
                }
                // stack 内查找
                if (item.type === 'stack' && item.items) {
                    const found = item.items.find(i => i.id === fieldId && typeof i.active === 'function');
                    if (found) return found;
                }
                // row 内查找
                if (item.type === 'row' && item.items) {
                    const found = item.items.find(i => i.id === fieldId && typeof i.active === 'function');
                    if (found) return found;
                }
            }
        }
    }
    return null;
}
