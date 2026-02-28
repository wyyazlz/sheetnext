/**
 * 条件格式XML构建模块
 * 负责将条件格式导出为Excel XML格式
 */
import { expandHex } from './helpers.js';

/**
 * 构建条件格式XML节点
 * @param {ConditionalFormat} cf - 条件格式管理器实例
 * @returns {Array|null} conditionalFormatting节点数组
 */
export function buildConditionalFormattingXml(cf) {
    if (!cf || cf.rules.length === 0) return null;
    return cf.toXmlObject();
}

/**
 * 构建dxf样式数组
 * @param {ConditionalFormat} cf - 条件格式管理器实例
 * @returns {Array} dxf节点数组
 */
export function buildDxfXml(cf) {
    if (!cf) return [];

    const dxfList = cf.getDxfList();
    return dxfList.map(dxf => buildSingleDxf(dxf));
}

/**
 * 构建单个dxf节点
 * @param {Object} dxf - 差异化样式对象
 * @returns {Object} dxf XML节点
 */
function buildSingleDxf(dxf) {
    const node = {};

    // 字体
    if (dxf.font) {
        const fontNode = {};
        if (dxf.font.bold) fontNode.b = '';
        if (dxf.font.italic) fontNode.i = '';
        if (dxf.font.strike) fontNode.strike = '';
        if (dxf.font.underline) {
            fontNode.u = dxf.font.underline === 'single' ? '' : { '_$val': dxf.font.underline };
        }
        if (dxf.font.color) {
            fontNode.color = { '_$rgb': expandHex(dxf.font.color) };
        }
        if (Object.keys(fontNode).length > 0) {
            node.font = fontNode;
        }
    }

    // 填充
    if (dxf.fill) {
        const fillNode = {};
        if (dxf.fill.type === 'pattern' || dxf.fill.fgColor) {
            const patternFill = {
                '_$patternType': dxf.fill.pattern || 'solid'
            };
            if (dxf.fill.fgColor) {
                const rgb = expandHex(dxf.fill.fgColor);
                patternFill.fgColor = { '_$rgb': rgb };
                patternFill.bgColor = { '_$rgb': rgb };
            }
            fillNode.patternFill = patternFill;
        }
        if (Object.keys(fillNode).length > 0) {
            node.fill = fillNode;
        }
    }

    // 边框
    if (dxf.border) {
        const borderNode = {};
        const sides = ['left', 'right', 'top', 'bottom'];
        for (const side of sides) {
            if (dxf.border[side]) {
                const sideNode = {};
                if (dxf.border[side].style) {
                    sideNode['_$style'] = dxf.border[side].style;
                }
                if (dxf.border[side].color) {
                    sideNode.color = { '_$rgb': expandHex(dxf.border[side].color) };
                }
                if (Object.keys(sideNode).length > 0) {
                    borderNode[side] = sideNode;
                }
            }
        }
        if (Object.keys(borderNode).length > 0) {
            node.border = borderNode;
        }
    }

    return node;
}

/**
 * 合并dxf到styles.xml
 * @param {Object} styleSheet - styles.xml的styleSheet节点
 * @param {Array} dxfArray - dxf数组
 */
export function mergeDxfToStyleSheet(styleSheet, dxfArray) {
    if (!dxfArray || dxfArray.length === 0) {
        return;
    }

    // 获取或创建dxfs节点
    if (!styleSheet.dxfs) {
        styleSheet.dxfs = { '_$count': '0' };
    }

    // 确保dxf是数组
    if (!styleSheet.dxfs.dxf) {
        styleSheet.dxfs.dxf = [];
    } else if (!Array.isArray(styleSheet.dxfs.dxf)) {
        styleSheet.dxfs.dxf = [styleSheet.dxfs.dxf];
    }

    // 添加新的dxf
    styleSheet.dxfs.dxf.push(...dxfArray);
    styleSheet.dxfs['_$count'] = String(styleSheet.dxfs.dxf.length);
}

export default {
    buildConditionalFormattingXml,
    buildDxfXml,
    mergeDxfToStyleSheet
};
