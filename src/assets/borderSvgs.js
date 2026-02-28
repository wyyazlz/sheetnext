/**
 * 边框 SVG 图标配置
 */

// 边框线条样式 - 使用 stroke-dasharray
export const BorderStyles = {
    hair: { dash: '1,2', width: 1 },
    dotted: { dash: '2,2', width: 1 },
    dashed: { dash: '4,2', width: 1 },
    dashDot: { dash: '6,2,2,2', width: 1 },
    dashDotDot: { dash: '6,2,2,2,2,2', width: 1 },
    slantDashDot: { dash: '6,2,2,2', width: 1 },
    thin: { dash: '', width: 1 },
    medium: { dash: '', width: 2 },
    mediumDashed: { dash: '6,3', width: 2 },
    mediumDashDot: { dash: '8,3,3,3', width: 2 },
    mediumDashDotDot: { dash: '8,3,3,3,3,3', width: 2 },
    thick: { dash: '', width: 3 },
    double: { dash: '', width: 1, double: true }
};

// 边框位置 - path 数据 (24x24 viewBox)
export const BorderPositions = {
    none: { paths: [], labelKey: 'action.style.border.position.none' },
    all: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true },
            { d: 'M12 3v18M3 12h18', stroke: true }
        ],
        labelKey: 'action.style.border.position.all'
    },
    outer: {
        paths: [{ d: 'M3 3h18v18H3z', stroke: true }],
        labelKey: 'action.style.border.position.outer'
    },
    inner: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true, light: true },
            { d: 'M12 3v18M3 12h18', stroke: true }
        ],
        labelKey: 'action.style.border.position.inner'
    },
    left: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true, light: true },
            { d: 'M3 3v18', stroke: true }
        ],
        labelKey: 'action.style.border.position.left'
    },
    right: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true, light: true },
            { d: 'M21 3v18', stroke: true }
        ],
        labelKey: 'action.style.border.position.right'
    },
    top: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true, light: true },
            { d: 'M3 3h18', stroke: true }
        ],
        labelKey: 'action.style.border.position.top'
    },
    bottom: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true, light: true },
            { d: 'M3 21h18', stroke: true }
        ],
        labelKey: 'action.style.border.position.bottom'
    },
    diagonalDown: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true, light: true },
            { d: 'M3 3l18 18', stroke: true }
        ],
        labelKey: 'action.style.border.position.diagonalDown'
    },
    diagonalUp: {
        paths: [
            { d: 'M3 3h18v18H3z', stroke: true, light: true },
            { d: 'M3 21l18-18', stroke: true }
        ],
        labelKey: 'action.style.border.position.diagonalUp'
    }
};

/**
 * 生成边框样式预览 SVG
 * @param {string} style - 样式名
 * @param {number} w - 宽度
 * @param {number} h - 高度
 */
export function getStyleSvg(style, w = 80, h = 20) {
    const s = BorderStyles[style];
    if (!s) return '';
    const y = h / 2;
    if (s.double) {
        return `<svg class="sn-svg" width="${w}" height="${h}"><line x1="4" y1="${y-2}" x2="${w-4}" y2="${y-2}" stroke="currentColor" stroke-width="1"/><line x1="4" y1="${y+2}" x2="${w-4}" y2="${y+2}" stroke="currentColor" stroke-width="1"/></svg>`;
    }
    return `<svg class="sn-svg" width="${w}" height="${h}"><line x1="4" y1="${y}" x2="${w-4}" y2="${y}" stroke="currentColor" stroke-width="${s.width}" stroke-dasharray="${s.dash}"/></svg>`;
}

/**
 * 生成边框位置图标 SVG
 * @param {string} pos - 位置名
 * @param {number} size - 尺寸
 */
export function getPosSvg(pos, size = 24) {
    const p = BorderPositions[pos];
    if (!p) return '';
    // 田字底纹
    const bg = `<path d="M3 3h18v18H3zM12 3v18M3 12h18" fill="none" stroke="#ddd" stroke-width="1"/>`;
    const paths = p.paths.map(item => {
        const color = item.light ? '#ccc' : 'currentColor';
        return `<path d="${item.d}" fill="none" stroke="${color}" stroke-width="1.6"/>`;
    }).join('');
    return `<svg class="sn-svg" width="${size}" height="${size}" viewBox="0 0 24 24">${bg}${paths}</svg>`;
}

export default { BorderStyles, BorderPositions, getStyleSvg, getPosSvg };
