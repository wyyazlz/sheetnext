/**
 * 菜单配置模块
 * 只保留右侧菜单配置
 */

/**
 * Initialize menu config.
 * @param {Object} SN - SheetNext instance.
 * @param {Object} [options={}] - Menu config options.
 * @param {function} [options.menuRight] - Callback `(defaultHTML: string) => string` to customize the top-right menu HTML.
 * @returns {Object} Menu config object.
 */
export function initDefaultMenuConfig(SN, options = {}) {
    // 默认右侧菜单 HTML
    const defaultRightLabel = SN.t('layout.menuRightMinimalToolbar');
    const defaultRightTitle = SN.t('layout.menuRightMinimalToolbarTitle');
    const defaultRightHTML = `
        <label class="sn-menu-minimal-toggle" title="${defaultRightTitle}">
            <input class="sn-menu-minimal-toggle-input" type="checkbox" data-role="menu-minimal-toolbar-toggle">
            <span class="sn-menu-minimal-toggle-text">${defaultRightLabel}</span>
        </label>
    `;

    let rightHTML = defaultRightHTML;
    if (typeof options.menuRight === 'function') {
        rightHTML = options.menuRight(defaultRightHTML);
    }

    return {
        right: { html: rightHTML }
    };
}
