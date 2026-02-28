// ==================== Cell XML构建 ====================
import { BUILTIN_NUMFMTS } from './constants.js';
import { expandHex } from './helpers.js';

function buildFontColorNode(font) {
    if (font.colorTheme !== undefined && font.colorTheme !== null && font.colorTheme !== '') {
        const colorNode = { _$theme: String(font.colorTheme) };
        if (font.colorTint !== undefined && font.colorTint !== null && font.colorTint !== '') {
            colorNode._$tint = String(font.colorTint);
        }
        return colorNode;
    }
    if (font.colorIndexed !== undefined && font.colorIndexed !== null && font.colorIndexed !== '') {
        return { _$indexed: String(font.colorIndexed) };
    }
    if (font.colorAuto !== undefined && font.colorAuto !== null && font.colorAuto !== '') {
        return { _$auto: String(font.colorAuto) };
    }
    if (font.color) {
        return { _$rgb: expandHex(font.color) };
    }
    return null;
}

function hasEffectiveFill(fill) {
    if (!fill || !fill.type || fill.type === 'none') return false;
    if (fill.type === 'pattern') {
        const pattern = fill.pattern || 'none';
        const hasColor = !!fill.fgColor || !!fill.bgColor;
        return pattern !== 'none' || hasColor;
    }
    return true;
}

function hasEffectiveBorder(border) {
    if (!border) return false;
    const hasSide = (side) => !!(side && (side.style || side.color));
    if (hasSide(border.top) || hasSide(border.right) || hasSide(border.bottom) || hasSide(border.left)) return true;
    if (!border.diagonal) return false;
    return hasSide(border.diagonal) && (
        border.diagonalDown === true
        || border.diagonalUp === true
        || (border.diagonalDown === undefined && border.diagonalUp === undefined)
    );
}

function buildBorderNode(border) {
    const borderObj = { top: "", right: "", bottom: "", left: "", diagonal: "", ...border };
    const ob = {};
    const sides = ['left', 'right', 'top', 'bottom', 'diagonal'];
    sides.forEach(side => {
        const sideStyle = borderObj[side];
        if (sideStyle) {
            ob[side] = {};
            if (sideStyle.style) ob[side]._$style = sideStyle.style;
            if (sideStyle.color) ob[side].color = { _$rgb: expandHex(sideStyle.color) };
            if (!Object.keys(ob[side]).length) ob[side] = "";
        } else {
            ob[side] = "";
        }
    });
    const drawDiagonalDown = borderObj.diagonal && (borderObj.diagonalDown === true || (borderObj.diagonalDown === undefined && borderObj.diagonalUp === undefined));
    const drawDiagonalUp = borderObj.diagonal && borderObj.diagonalUp === true;
    if (drawDiagonalDown) ob._$diagonalDown = '1';
    if (drawDiagonalUp) ob._$diagonalUp = '1';
    return { key: JSON.stringify(borderObj), node: ob };
}

/**
 * 检查单元格是否在超级表的数据区域（非表头）
 * @param {Cell} cell
 * @returns {boolean}
 */
function isInTableDataArea(cell) {
    const sheet = cell.row?.sheet;
    if (!sheet?.Table || sheet.Table.size === 0) return false;

    let inDataArea = false;
    sheet.Table.forEach(table => {
        // 只检查数据区域，不包括表头（表头可以有自己的样式）
        if (table.isDataCell(cell.row.rIndex, cell.cIndex)) {
            inDataArea = true;
        }
    });
    return inDataArea;
}

/**
 * 构建通用样式 XF 对象（供行/列/单元格使用）
 * @param {Object} style - 样式对象 { font, fill, alignment, border, numFmt, protection }
 * @param {Object} SN - SheetNext 实例
 * @param {Object} options - 可选配置 { skipFillForTable: boolean }
 * @returns {Object} xfObj
 */
export function buildStyleXfFromObject(style, SN, options = {}) {
    if (!style || Object.keys(style).length === 0) return null;

    const xfObj = {
        _$numFmtId: 0,
        _$fontId: 0,
        _$fillId: 0,
        _$borderId: 0,
        _$xfId: 0
    };

    // 对齐方式构建
    if (style.alignment) {
        xfObj.alignment = {};
        if (style.alignment.horizontal) xfObj.alignment._$horizontal = style.alignment.horizontal;
        if (style.alignment.vertical) xfObj.alignment._$vertical = style.alignment.vertical;
        if (style.alignment.wrapText) xfObj.alignment._$wrapText = '1';
        if (style.alignment.shrinkToFit) xfObj.alignment._$shrinkToFit = '1';
        if (style.alignment.indent) xfObj.alignment._$indent = style.alignment.indent;
        if (style.alignment.textRotation) xfObj.alignment._$textRotation = style.alignment.textRotation;
    }

    // FONT构建
    if (style.font) {
        const font = style.font;
        const key = JSON.stringify(font);
        if (SN.buildStyle.fonts[key]) {
            xfObj._$fontId = SN.buildStyle.fonts[key];
        } else {
            xfObj._$fontId = SN.buildStyle.fonts[key] = SN.buildStyle.fontC++;
            const ob = {};
            if (font.bold) ob.b = "";
            if (font.italic) ob.i = "";
            if (font.strike) ob.strike = "";
            if (font.underline) ob.u = font.underline === 'single' ? "" : { _$val: font.underline };
            if (font.size) ob.sz = { _$val: font.size };
            const fontColorNode = buildFontColorNode(font);
            if (fontColorNode) ob.color = fontColorNode;
            if (font.name) ob.name = { _$val: font.name };
            if (font.charset) ob.charset = { _$val: font.charset };
            if (font.outline) ob.outline = "";
            if (font.vertAlign) ob.vertAlign = { _$val: font.vertAlign };
            SN._Xml.styleSheet.fonts.font.push(ob);
        }
        xfObj._$applyFont = "1";
    }

    // Fill构建
    if (!options.skipFillForTable && hasEffectiveFill(style.fill)) {
        const fill = style.fill;
        const key = JSON.stringify(fill);
        if (SN.buildStyle.fills[key]) {
            xfObj._$fillId = SN.buildStyle.fills[key];
        } else {
            xfObj._$fillId = SN.buildStyle.fills[key] = SN.buildStyle.fillC++;
            const ob = {};
            if (fill.type === 'pattern') {
                ob.patternFill = { _$patternType: fill.pattern || 'solid' };
                if (fill.fgColor) ob.patternFill.fgColor = { _$rgb: expandHex(fill.fgColor) };
                if (fill.bgColor) ob.patternFill.bgColor = { _$rgb: expandHex(fill.bgColor) };
            } else if (fill.type === 'gradient') {
                ob.gradientFill = {};
                if (fill.gradientType === 'path') ob.gradientFill._$type = 'path';
                if (fill.degree) ob.gradientFill._$degree = fill.degree;
                if (fill.left) ob.gradientFill._$left = fill.left;
                if (fill.right) ob.gradientFill._$right = fill.right;
                if (fill.top) ob.gradientFill._$top = fill.top;
                if (fill.bottom) ob.gradientFill._$bottom = fill.bottom;

                if (fill.stops && Array.isArray(fill.stops)) {
                    ob.gradientFill.stop = fill.stops.map(stop => {
                        const stopObj = {};
                        stopObj._$position = stop.position !== undefined ? stop.position : 0;
                        if (stop.color) stopObj.color = { _$rgb: expandHex(stop.color) };
                        return stopObj;
                    });
                } else {
                    ob.gradientFill.stop = [
                        { _$position: 0, color: { _$rgb: "FFFFFFFF" } },
                        { _$position: 1, color: { _$rgb: "FF000000" } }
                    ];
                }
            }
            SN._Xml.styleSheet.fills.fill.push(ob);
        }
        xfObj._$applyFill = "1";
    }

    // BORDER构建
    if (hasEffectiveBorder(style.border)) {
        const borderNode = buildBorderNode(style.border);
        const key = borderNode.key;
        if (SN.buildStyle.borders[key]) {
            xfObj._$borderId = SN.buildStyle.borders[key];
        } else {
            xfObj._$borderId = SN.buildStyle.borders[key] = SN.buildStyle.borderC++;
            SN._Xml.styleSheet.borders.border.push(borderNode.node);
        }
        xfObj._$applyBorder = "1";
    }

    // numFmtId构建
    if (style.numFmt) {
        const numFmt = style.numFmt;
        const key = JSON.stringify(numFmt);
        if (SN.buildStyle.numFmts[key]) {
            xfObj._$numFmtId = SN.buildStyle.numFmts[key];
        } else {
            const builtInId = Object.entries(BUILTIN_NUMFMTS).find(([_, format]) => format === numFmt)?.[0];
            if (builtInId) {
                xfObj._$numFmtId = builtInId;
            } else {
                xfObj._$numFmtId = SN.buildStyle.numFmtC++;
                SN._Xml.styleSheet.numFmts.numFmt.push({
                    _$numFmtId: xfObj._$numFmtId,
                    _$formatCode: numFmt
                });
            }
            SN.buildStyle.numFmts[key] = xfObj._$numFmtId;
        }
        xfObj._$applyNumberFormat = "1";
    }

    // protection 构建
    if (style.protection) {
        const protection = {};
        if (style.protection.locked === false) protection._$locked = '0';
        if (style.protection.hidden === true) protection._$hidden = '1';
        if (Object.keys(protection).length > 0) {
            xfObj.protection = protection;
            xfObj._$applyProtection = "1";
        }
    }

    return xfObj;
}

/**
 * 获取样式索引（供行/列使用）
 * @param {Object} style - 样式对象
 * @param {Object} SN - SheetNext 实例
 * @returns {number|null} 样式索引
 */
export function getStyleIndex(style, SN) {
    const xfObj = buildStyleXfFromObject(style, SN);
    if (!xfObj) return null;

    const xfKey = JSON.stringify(xfObj);
    if (SN.buildStyle.xf[xfKey] != undefined) {
        return SN.buildStyle.xf[xfKey];
    } else {
        const index = SN.buildStyle.xf[xfKey] = SN.buildStyle.xfC++;
        SN._Xml.styleSheet.cellXfs.xf.push(xfObj);
        return index;
    }
}

// 构建样式对象
export function buildStyleXf(cell) {
    const xfObj = {
        _$numFmtId: 0,
        _$fontId: 0,
        _$fillId: 0,
        _$borderId: 0,
        _$xfId: 0
    };
    const SN = cell._SN;

    // 对齐方式构建
    if (cell._style.alignment) {
        xfObj.alignment = {};
        if (cell.alignment.horizontal) xfObj.alignment._$horizontal = cell.alignment.horizontal;
        if (cell.alignment.vertical) xfObj.alignment._$vertical = cell.alignment.vertical;
        if (cell.alignment.wrapText) xfObj.alignment._$wrapText = '1';
        if (cell.alignment.shrinkToFit) xfObj.alignment._$shrinkToFit = '1';
        if (cell.alignment.indent) xfObj.alignment._$indent = cell.alignment.indent;
        if (cell.alignment.textRotation) xfObj.alignment._$textRotation = cell.alignment.textRotation;
    }

    // FONT构建
    if (cell._style.font) {
        const font = cell.font;
        const key = JSON.stringify(font);
        if (SN.buildStyle.fonts[key]) {
            xfObj._$fontId = SN.buildStyle.fonts[key];
        } else {
            xfObj._$fontId = SN.buildStyle.fonts[key] = SN.buildStyle.fontC++;
            const ob = {};
            if (font.bold) ob.b = "";
            if (font.italic) ob.i = "";
            if (font.strike) ob.strike = "";
            if (font.underline) ob.u = font.underline === 'single' ? "" : { _$val: font.underline };
            if (font.size) ob.sz = { _$val: font.size };
            const fontColorNode = buildFontColorNode(font);
            if (fontColorNode) ob.color = fontColorNode;
            if (font.name) ob.name = { _$val: font.name };
            if (font.charset) ob.charset = { _$val: font.charset };
            if (font.outline) ob.outline = "";
            if (font.vertAlign) ob.vertAlign = { _$val: font.vertAlign };
            SN._Xml.styleSheet.fonts.font.push(ob);
        }
        xfObj._$applyFont = "1";
    }

    // Fill构建（超级表数据区域内的单元格跳过 fill，让超级表样式接管）
    const skipFillForTable = isInTableDataArea(cell);
    if (!skipFillForTable && hasEffectiveFill(cell._style?.fill)) {
        const fill = cell.fill;
        const key = JSON.stringify(fill);
        if (SN.buildStyle.fills[key]) {
            xfObj._$fillId = SN.buildStyle.fills[key];
        } else {
            xfObj._$fillId = SN.buildStyle.fills[key] = SN.buildStyle.fillC++;
            const ob = {};
            if (fill.type === 'pattern') {
                ob.patternFill = { _$patternType: fill.pattern || 'solid' };
                if (fill.fgColor) ob.patternFill.fgColor = { _$rgb: expandHex(fill.fgColor) };
                if (fill.bgColor) ob.patternFill.bgColor = { _$rgb: expandHex(fill.bgColor) };
            } else if (fill.type === 'gradient') {
                ob.gradientFill = {};
                if (fill.gradientType === 'path') ob.gradientFill._$type = 'path';
                if (fill.degree) ob.gradientFill._$degree = fill.degree;
                if (fill.left) ob.gradientFill._$left = fill.left;
                if (fill.right) ob.gradientFill._$right = fill.right;
                if (fill.top) ob.gradientFill._$top = fill.top;
                if (fill.bottom) ob.gradientFill._$bottom = fill.bottom;

                if (fill.stops && Array.isArray(fill.stops)) {
                    ob.gradientFill.stop = fill.stops.map(stop => {
                        const stopObj = {};
                        stopObj._$position = stop.position !== undefined ? stop.position : 0;
                        if (stop.color) stopObj.color = { _$rgb: expandHex(stop.color) };
                        return stopObj;
                    });
                } else {
                    ob.gradientFill.stop = [
                        { _$position: 0, color: { _$rgb: "FFFFFFFF" } },
                        { _$position: 1, color: { _$rgb: "FF000000" } }
                    ];
                }
            }
            SN._Xml.styleSheet.fills.fill.push(ob);
        }
        xfObj._$applyFill = "1";
    }

    // BORDER构建
    if (hasEffectiveBorder(cell._style?.border)) {
        const borderNode = buildBorderNode(cell.border);
        const key = borderNode.key;
        if (SN.buildStyle.borders[key]) {
            xfObj._$borderId = SN.buildStyle.borders[key];
        } else {
            xfObj._$borderId = SN.buildStyle.borders[key] = SN.buildStyle.borderC++;
            SN._Xml.styleSheet.borders.border.push(borderNode.node);
        }
        xfObj._$applyBorder = "1";
    }

    // numFmtId构建
    if (cell._style.numFmt) {
        const numFmt = cell.numFmt;
        const key = JSON.stringify(numFmt);
        if (SN.buildStyle.numFmts[key]) {
            xfObj._$numFmtId = SN.buildStyle.numFmts[key];
        } else {
            const builtInId = Object.entries(BUILTIN_NUMFMTS).find(([_, format]) => format === numFmt)?.[0];
            if (builtInId) {
                xfObj._$numFmtId = builtInId;
            } else {
                xfObj._$numFmtId = SN.buildStyle.numFmtC++;
                SN._Xml.styleSheet.numFmts.numFmt.push({
                    _$numFmtId: xfObj._$numFmtId,
                    _$formatCode: numFmt
                });
            }
            SN.buildStyle.numFmts[key] = xfObj._$numFmtId;
        }
        xfObj._$applyNumberFormat = "1";
    }

    // protection 构建
    if (cell._style?.protection) {
        const protection = {};
        if (cell._style.protection.locked === false) protection._$locked = '0';
        if (cell._style.protection.hidden === true) protection._$hidden = '1';
        if (Object.keys(protection).length > 0) {
            xfObj.protection = protection;
            xfObj._$applyProtection = "1";
        }
    }

    return xfObj;
}

// 构建单元格XML对象
export function buildCellXml(cell) {
    let obj = {
        "v": "",
        "_$r": `${cell._SN.Utils.numToChar(cell.cIndex)}${cell.row.rIndex + 1}`
    };
    const SN = cell._SN;

    // 类型构建
    if (cell.isFormula) {
        obj = { f: cell.editVal.substring(1), ...obj };
    } else if (cell.type === 'boolean') {
        obj['_$t'] = 'b';
        obj.v = cell._editVal ? '1' : '0';
    } else if (['time', 'dateTime', 'date'].includes(cell.type)) {
        obj.v = cell.calcVal;
    } else if (cell.type === 'error') {
        obj["_$t"] = 'e';
        obj.v = cell._showVal;
    } else if (isNaN(cell._editVal)) {
        obj["_$t"] = 's';
        const sst = SN._Xml.sst;
        if (cell._richText) {
            // 富文本：构建 r 数组结构
            const key = JSON.stringify(cell._richText);
            if (SN._Xml.sharedStringsObj[key] === undefined) {
                SN._Xml.sharedStringsObj[key] = Object.keys(SN._Xml.sharedStringsObj).length;
                sst['_$uniqueCount']++;
                const rArr = cell._richText.map(run => {
                    const rf = run.font;
                    // rPr 必须在 t 之前（Excel 要求的 XML 顺序）
                    const entry = {};
                    if (rf && Object.keys(rf).length) {
                        // rPr 子元素顺序遵循 OOXML CT_RPrElt 规范
                        const rPr = {};
                        // rFont 和 sz 是必需的，缺失时 fallback 到单元格字体
                        const cellFont = cell.font;
                        rPr.rFont = { _$val: rf.name ?? cellFont.name ?? '宋体' };
                        if (rf.charset ?? cellFont.charset) rPr.charset = { _$val: rf.charset ?? cellFont.charset };
                        if (rf.bold ?? cellFont.bold) rPr.b = "";
                        if (rf.italic ?? cellFont.italic) rPr.i = "";
                        if (rf.strike ?? cellFont.strike) rPr.strike = "";
                        if (rf.color) {
                            rPr.color = { _$rgb: 'FF' + rf.color.replace('#', '') };
                        }
                        rPr.sz = { _$val: rf.size ?? cellFont.size ?? '11' };
                        const ul = rf.underline ?? cellFont.underline;
                        if (ul) {
                            rPr.u = ul === 'single' ? "" : { _$val: ul };
                        }
                        if (rf.vertAlign) rPr.vertAlign = { _$val: rf.vertAlign };
                        entry.rPr = rPr;
                    }
                    entry.t = run.text;
                    return entry;
                });
                sst.si.push({ r: rArr });
            }
            sst['_$count']++;
            obj.v = SN._Xml.sharedStringsObj[key];
        } else {
            // 纯文本
            const key = cell.editVal;
            if (SN._Xml.sharedStringsObj[key] === undefined) {
                SN._Xml.sharedStringsObj[key] = Object.keys(SN._Xml.sharedStringsObj).length;
                sst['_$uniqueCount']++;
                sst.si.push({ t: key });
            }
            sst['_$count']++;
            obj.v = SN._Xml.sharedStringsObj[key];
        }
    } else {
        obj.v = cell.editVal;
    }

    // 样式构建
    const hasLineBreak = typeof cell.editVal === 'string' && cell.editVal.includes('\n');
    if (Object.keys(cell.style).length == 0 && !hasLineBreak) return obj;

    const xfObj = buildStyleXf(cell);
    // 含换行符时强制设置 wrapText，Excel 需要此属性才能显示换行
    if (hasLineBreak) {
        if (!xfObj.alignment) xfObj.alignment = {};
        xfObj.alignment._$wrapText = '1';
    }
    const xfKey = JSON.stringify(xfObj);

    if (SN.buildStyle.xf[xfKey] != undefined) {
        obj._$s = SN.buildStyle.xf[xfKey];
    } else {
        obj._$s = SN.buildStyle.xf[xfKey] = SN.buildStyle.xfC++;
        SN._Xml.styleSheet.cellXfs.xf.push(xfObj);
    }

    return obj;
}
