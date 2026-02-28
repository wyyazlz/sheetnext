/**
 * 条件格式渲染模块
 * 负责在Canvas上渲染条件格式效果（数据条、图标）
 * 注意：背景色和字体色已在 rBaseData 中预计算并合并渲染
 */
import { getIconCanvas } from '../CF/icons.js';
import { lightenColor } from '../CF/helpers.js';
import { ICON_SETS } from '../CF/rules/IconSetRule.js';

/**
 * 渲染条件格式（仅图标）
 * 数据条已在 rBaseData 中渲染，避免覆盖文字
 */
export function rConditionalFormats(sheet) {
    const cf = sheet.CF;
    if (!cf || cf.rules.length === 0) return;

    const ctx = this.bc;
    const pixelRatio = this.pixelRatio;
    const cellPositions = this._cfCellPositions;

    // 如果没有缓存的位置信息，跳过
    if (!cellPositions) return;

    // 遍历已缓存的单元格位置（已在rBaseData中计算）
    cellPositions.forEach((pos, key) => {
        const { x, y, w, h, cellInfo, cfFormat } = pos;

        if (cfFormat) {
            // 渲染图标（在文字之后，显示在单元格左侧）
            if (cfFormat.icon) {
                _rIcon(ctx, x, y, w, h, cfFormat.icon, cellInfo, pixelRatio);
            }
        }
    });
}

/**
 * 渲染数据条（导出供 CellRenderer 使用）
 */
export function rDataBar(x, y, w, h, dataBar) {
    _rDataBar(this.bc, x, y, w, h, dataBar, this.pixelRatio);
}

/**
 * 内部数据条渲染函数
 */
function _rDataBar(ctx, x, y, w, h, dataBar, pixelRatio) {
    const snap = (value) => Math.round(value * pixelRatio) / pixelRatio;
    const lineWidth = 1 / pixelRatio;
    const { percent, color, gradient, borderColor, axisPosition, axisColor, isNegative, showValue } = dataBar;

    const padding = 2;
    const barH = h - padding * 2;
    const barY = y + padding;

    // 计算条的位置和宽度
    let barX, barW;

    if (axisPosition !== null && axisPosition !== undefined) {
        // 有正负值：轴线在中间
        const axisX = x + w * axisPosition;

        // 绘制轴线
        ctx.strokeStyle = axisColor || '#000000';
        ctx.lineWidth = lineWidth;
        ctx.beginPath();
        const axisXLine = snap(axisX);
        ctx.moveTo(axisXLine, snap(y));
        ctx.lineTo(axisXLine, snap(y + h));
        ctx.stroke();

        if (percent >= 0) {
            // 正值：从轴线向右
            barX = axisX;
            barW = (w - (axisX - x)) * Math.abs(percent);
        } else {
            // 负值：从轴线向左
            barW = (axisX - x) * Math.abs(percent);
            barX = axisX - barW;
        }
    } else {
        // 全正或全负：从左到右
        barX = x + padding;
        barW = (w - padding * 2) * Math.abs(percent);
    }

    barW = Math.max(0, barW);

    // 绘制数据条
    if (barW > 0) {
        if (gradient) {
            // 渐变填充：左边深色 → 右边浅色（淡出效果）
            const grad = ctx.createLinearGradient(barX, barY, barX + barW, barY);
            grad.addColorStop(0, color);
            grad.addColorStop(1, lightenColor(color, 0.65));
            ctx.fillStyle = grad;
        } else {
            // 纯色填充
            ctx.fillStyle = color;
        }

        ctx.fillRect(barX, barY, barW, barH);

        // 边框
        if (borderColor) {
            ctx.strokeStyle = borderColor;
            ctx.lineWidth = lineWidth;
            const x1 = snap(barX);
            const y1 = snap(barY);
            const x2 = snap(barX + barW);
            const y2 = snap(barY + barH);
            ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
        }
    }
}

/**
 * 渲染图标
 */
function _rIcon(ctx, x, y, w, h, iconConfig, cell, pixelRatio) {
    const snap = (value) => Math.round(value * pixelRatio) / pixelRatio;
    const { set, index, showValue } = iconConfig;

    // 获取图标集配置
    const setConfig = ICON_SETS[set];
    if (!setConfig || !setConfig.icons[index]) return;

    const iconName = setConfig.icons[index];
    const iconSize = Math.min(16, h - 4);

    // 按pixelRatio倍率绘制，避免高DPI模糊
    const drawSize = Math.round(iconSize * pixelRatio);
    const iconCanvas = getIconCanvas(iconName, drawSize);
    if (!iconCanvas) return;

    // 绘制图标（对齐像素避免模糊）
    const iconX = snap(Math.floor(x + 2));
    const iconY = snap(Math.floor(y + (h - iconSize) / 2));

    ctx.drawImage(iconCanvas, iconX, iconY, iconSize, iconSize);

    // 如果不显示值，需要在后续渲染中跳过文本
    // 这里通过修改cell的临时属性来标记
    if (!showValue) {
        cell._cfHideValue = true;
    }
}

/**
 * 获取条件格式应用后的样式（用于文本渲染）
 * @param {Cell} cell - 单元格
 * @param {Object} format - 条件格式结果
 * @returns {Object} 合并后的样式
 */
export function getCFMergedStyle(cell, format) {
    if (!format) return cell.style;

    const baseStyle = cell.style || {};
    const result = { ...baseStyle };

    // 合并字体样式
    if (format.font) {
        result.font = { ...baseStyle.font, ...format.font };
    }

    // 合并边框样式
    if (format.border) {
        result.border = { ...baseStyle.border, ...format.border };
    }

    // 图标偏移（文本需要右移）
    if (format.icon) {
        result._iconOffset = format.icon.showValue !== false ? 18 : 0;
        result._hideValue = !format.icon.showValue;
    }

    return result;
}

export default {
    rConditionalFormats,
    rDataBar,
    getCFMergedStyle
};
