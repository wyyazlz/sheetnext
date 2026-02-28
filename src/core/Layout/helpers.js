/**
 * Layout 辅助方法
 * 包含菜单验证、生成、颜色组件初始化等辅助功能
 */

// ============================= 菜单验证 =============================

/**
 * 递归验证菜单项（包含子菜单）
 * @param {Array} items - 菜单项数组
 * @returns {boolean} 是否有效
 */
function validateMenuItems(items) {
    if (!Array.isArray(items)) return false;

    for (const item of items) {
        // divider 项只需要有 divider 属性
        if (item.divider) continue;

        // 其他项必须有 labelKey 或 text
        if (!item.labelKey && !item.text && !item.label) return false;

        // 如果有 handler，必须是函数
        if (item.handler && typeof item.handler !== 'function') return false;

        // 如果有 customHTML，必须是字符串
        if (item.customHTML && typeof item.customHTML !== 'string') {
            return false;
        }

        // 如果有 submenu，递归验证子菜单
        if (item.submenu) {
            if (!validateMenuItems(item.submenu)) return false;
        }
    }

    return true;
}

/**
 * 验证菜单列表结构是否正确（支持多级菜单）
 * @param {Array} menuList - 菜单列表
 * @returns {boolean} 是否有效
 */
export function validateMenuList(menuList) {
    if (!Array.isArray(menuList)) return false;

    for (const menu of menuList) {
        // 每个菜单必须有 labelKey 或 text，以及 items
        if ((!menu.labelKey && !menu.text && !menu.label) || !Array.isArray(menu.items)) return false;

        // 验证每个 item（支持递归验证子菜单）
        if (!validateMenuItems(menu.items)) return false;
    }

    return true;
}

// ============================= 菜单生成 =============================

/**
 * 生成唯一的 handler ID
 * @param {Object} context - 包含计数器的上下文对象
 * @returns {string} handler ID
 */
export function generateHandlerId(context) {
    if (!context._handlerCounter) context._handlerCounter = 0;
    return `menuHandler_${context._handlerCounter++}`;
}

/**
 * 挂载 handler 到 SN.Action 上并返回方法名
 * @param {Function} handler - 处理函数
 * @param {Object} SN - SheetNext 实例
 * @param {Object} context - 包含计数器的上下文对象
 * @returns {string} handler 方法名
 */
export function mountHandler(handler, SN, context) {
    if (typeof handler !== 'function') return '';

    const handlerId = generateHandlerId(context);
    SN.Action[handlerId] = handler;
    return handlerId;
}

/**
 * 翻译文本
 * @param {Object} SN - SheetNext 实例
 * @param {string} text - 原始文本
 * @returns {string} 翻译后文本
 */
function _tr(SN, text) {
    if (typeof text !== 'string') return '';
    return SN.t(text);
}

function _resolveText(SN, item) {
    if (!item) return '';
    if (typeof item.labelKey === 'string' && item.labelKey) return SN.t(item.labelKey);
    return _tr(SN, item.text || item.label);
}

/**
 * 递归生成菜单项 HTML（支持多级子菜单和自定义HTML）
 * @param {Array} items - 菜单项数组
 * @param {string} namespace - 命名空间
 * @param {Object} SN - SheetNext 实例
 * @param {Object} context - 包含计数器的上下文对象
 * @returns {string} HTML 字符串
 */
function generateMenuItemsHTML(items, namespace, SN, context) {
    let html = '';

    items.forEach(item => {
        const label = _resolveText(SN, item);
        const tip = item.tip ? _tr(SN, item.tip) : '';
        if (item.divider) {
            // 分隔线
            html += '<li class="sn-dropdown-divider"></li>';
        } else if (item.customHTML) {
            // 自定义 HTML 内容（字符串）
            html += `<li>
                <span>${label}${tip ? `<span class="sn-float-end sn-color-999">${tip}</span>` : ''}</span>
                <div class="sn-dropdown-submenu">
                    ${item.customHTML}
                </div>
            </li>`;
        } else if (item.submenu && Array.isArray(item.submenu) && item.submenu.length > 0) {
            // 有子菜单的项
            html += `<li>
                <span>${label}${tip ? `<span class="sn-float-end sn-color-999">${tip}</span>` : ''}</span>
                <ul class="sn-dropdown-submenu">`;

            // 递归生成子菜单
            html += generateMenuItemsHTML(item.submenu, namespace, SN, context);

            html += '</ul></li>';
        } else if (item.disabled) {
            // 禁用的项
            html += `<li data-ds="true"><span>${label}${tip ? `<span class="sn-float-end sn-color-999">${tip}</span>` : ''}</span></li>`;
        } else {
            // 普通项（可能有 handler）
            let onclick = '';
            if (item.handler) {
                const handlerName = mountHandler(item.handler, SN, context);
                onclick = `onclick="${namespace}.Action.${handlerName}()"`;
            }
            html += `<li ${onclick}><span>${label}${tip ? `<span class="sn-float-end sn-color-999">${tip}</span>` : ''}</span></li>`;
        }
    });

    return html;
}

/**
 * 生成菜单列表 HTML（支持多级菜单）
 * @param {Array} menuList - 菜单列表配置
 * @param {string} namespace - 命名空间
 * @param {Object} SN - SheetNext 实例
 * @param {Object} context - 包含计数器的上下文对象
 * @returns {string} HTML 字符串
 */
export function generateMenuListHTML(menuList, namespace, SN, context) {
    let html = '';

    menuList.forEach(menu => {
        const label = _resolveText(SN, menu);
        html += `<div class="sn-dropdown">
            <div class="head-btn sn-dropdown-toggle">${label}</div>
            <ul class="sn-dropdown-menu"${menu.minWidth ? ` style="min-width: ${menu.minWidth};"` : ''}>`;

        // 使用递归函数生成菜单项
        html += generateMenuItemsHTML(menu.items, namespace, SN, context);

        html += '</ul></div>';
    });

    return html;
}

// ============================= 颜色组件 =============================

// 颜色配置
const COLORS = {
    common: [
        "#FFFFFF", "#000000", "#485368", "#2972F4", "#00A3F5", "#319B62", "#DE3C36", "#F88825", "#F5C400", "#9A38D7",
        "#F2F2F2", "#7F7F7F", "#F3F5F7", "#E5EFFF", "#E5F6FF", "#EAFAF1", "#FFE9E8", "#FFF3EB", "#FFF9E3", "#FDEBFF",
        "#D8D8D8", "#595959", "#C5CAD3", "#C7DCFF", "#C7ECFF", "#C3EAD5", "#FFC9C7", "#FFDCC4", "#FFEEAD", "#F2C7FF",
        "#BFBFBF", "#3F3F3F", "#808B9E", "#99BEFF", "#99DDFF", "#98D7B6", "#FF9C99", "#FFBA84", "#FFE270", "#D58EFF",
        "#A5A5A5", "#262626", "#353B45", "#1450B8", "#1274A5", "#277C4F", "#9E1E1A", "#B86014", "#A38200", "#5E2281",
        "#939393", "#0D0D0D", "#24272E", "#0C306E", "#0A415C", "#184E32", "#58110E", "#5C300A", "#665200", "#3B1551"
    ],
    standard: ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"]
};

/**
 * 生成颜色选择器 HTML
 * @returns {string} 颜色选择器 HTML
 */
export function getColorSelectHTML(SN) {
    const noColorText = SN.t('color.noColor');
    const themeColorsText = SN.t('color.themeColors');
    const standardColorsText = SN.t('color.standardColors');
    const moreColorsText = SN.t('color.moreColors');

    let html = `
        <div class="sn-color-section">
            <div class="sn-color-no-btn">${noColorText}</div>
        </div>
        <div class="sn-color-section">
            <div class="sn-color-label">${themeColorsText}</div>
            <div class="sn-color-grid">`;

    COLORS.common.forEach(c => {
        html += `<div class="sn-color-item" data-color="${c}" style="background:${c}"></div>`;
    });

    html += `</div>
        </div>
        <div class="sn-color-section">
            <div class="sn-color-label">${standardColorsText}</div>
            <div class="sn-color-grid">`;

    COLORS.standard.forEach(c => {
        html += `<div class="sn-color-item" data-color="${c}" style="background:${c}"></div>`;
    });

    html += `</div>
        </div>
        <div class="sn-color-section sn-color-more">
            <div class="sn-color-label">${moreColorsText}</div>
            <input type="color" class="sn-color-input" value="#336699">
        </div>`;

    return html;
}

// ============================= 拖拽辅助 =============================

/**
 * 计算拖拽边界
 * @param {DOMRect} mainRect - 主容器的边界矩形
 * @param {DOMRect} opRect - 操作容器的边界矩形
 * @param {DOMRect} chatRect - 聊天窗口的边界矩形
 * @returns {Object} 边界对象 {minX, minY, maxX, maxY}
 */
export function calculateDragBounds(mainRect, opRect, chatRect) {
    // 计算 snOp 相对于 snMain 的偏移
    const opOffsetX = opRect.left - mainRect.left;
    const opOffsetY = opRect.top - mainRect.top;

    // 计算在 snOp 坐标系下的边界（相对于 snMain 的边界）
    return {
        minX: -opOffsetX,
        minY: -opOffsetY,
        maxX: mainRect.width - chatRect.width - opOffsetX,
        maxY: mainRect.height - chatRect.height - opOffsetY,
        opOffsetX,
        opOffsetY
    };
}

/**
 * 限制坐标在边界内
 * @param {number} x - X 坐标
 * @param {number} y - Y 坐标
 * @param {Object} bounds - 边界对象
 * @returns {Object} 限制后的坐标 {x, y}
 */
export function constrainToBounds(x, y, bounds) {
    let constrainedX = x;
    let constrainedY = y;

    if (constrainedX < bounds.minX) constrainedX = bounds.minX;
    if (constrainedY < bounds.minY) constrainedY = bounds.minY;
    if (constrainedX > bounds.maxX) constrainedX = bounds.maxX;
    if (constrainedY > bounds.maxY) constrainedY = bounds.maxY;

    return { x: constrainedX, y: constrainedY };
}

// ============================= Viewport 设置 =============================

/**
 * 设置 viewport meta 标签
 */
export function setupViewport() {
    let viewport = document.querySelector('meta[name="viewport"]');
    if (!viewport) {
        viewport = document.createElement('meta');
        viewport.name = 'viewport';
        document.head.appendChild(viewport);
    }
    viewport.content = 'width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no';
}
