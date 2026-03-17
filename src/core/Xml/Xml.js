// ==================== 导入模块 ====================

import snXml from './template.js';
import { getColor, getColorByIndex as _getColorByIndex, getThemeColor as _getThemeColor } from './helpers.js';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { BUILTIN_NUMFMTS } from '../Cell/constants.js';

// ==================== Xml 类 ====================

/**
 * Xml Parsing and Builder
 * @title XML Parsing
 * @class
 */
export default class Xml {
    /**
     * Static XML parser (all examples shared)
     * @type {XMLParser}
     */
    static parser = new XMLParser({
        attributeNamePrefix: "_$",
        ignoreAttributes: false,
        parseAttributeValue: false,
        trimValues: false,
        parseTagValue: false,
        isArray: (tagName, jPath, isLeafNode, isAttribute) => {
            const paths = [
                'worksheet.sheetData.row',
                'workbook.sheets.sheet',
                'worksheet.cols.col',
                'worksheet.sheetViews.sheetView',
                'worksheet.mergeCells.mergeCell',
                'worksheet.hyperlinks.hyperlink',
                'worksheet.dataValidations.dataValidation',
                'worksheet.AutoFilter.filterColumn',
                'worksheet.tableParts.tablePart',
                'styleSheet.numFmts.numFmt',
                'Relationships.Relationship',
                'sst.si',
                // 超级表相关
                'table.tableColumns.tableColumn',
                // 透视表相关
                'workbook.pivotCaches.pivotCache',
                'pivotCacheDefinition.cacheFields.cacheField',
                'pivotCacheDefinition.cacheFields.cacheField.sharedItems.s',
                'pivotCacheDefinition.cacheFields.cacheField.sharedItems.n',
                'pivotCacheDefinition.cacheFields.cacheField.sharedItems.d',
                'pivotCacheDefinition.cacheFields.cacheField.sharedItems.b',
                'pivotCacheDefinition.cacheFields.cacheField.sharedItems.m',
                'pivotCacheRecords.r',
                'pivotCacheRecords.r.x',
                'pivotCacheRecords.r.n',
                'pivotCacheRecords.r.s',
                'pivotCacheRecords.r.d',
                'pivotCacheRecords.r.b',
                'pivotCacheRecords.r.m',
                'pivotTableDefinition.pivotFields.pivotField',
                'pivotTableDefinition.pivotFields.pivotField.items.item',
                'pivotTableDefinition.rowFields.field',
                'pivotTableDefinition.colFields.field',
                'pivotTableDefinition.dataFields.dataField',
                'pivotTableDefinition.pageFields.pageField',
                'pivotTableDefinition.formats.format'
            ];
            return paths.includes(jPath);
        }
    });

    /**
     * Static XML Builder (all examples shared)
     * @type {XMLBuilder}
     */
    static builder = new XMLBuilder({
        attributeNamePrefix: "_$",
        ignoreAttributes: false,
        suppressEmptyNode: true
    });

    /**
     * @param {Object} SN - SheetNext Main Instance
     * @param {Object} [obj] - XML Object
     */
    constructor(SN, obj) {
        /**
         * SheetNext Main Instance
         * @type {Object}
         */
        this.SN = SN
        /**
         * Tool Method Collection
         * @type {Utils}
         */
        this.Utils = SN.Utils
        /**
         * XML Object
         * @type {Object}
         */
        this.obj = obj ?? JSON.parse(JSON.stringify(snXml));
        /**
         * Share String to Heavy Object
         * @type {Object}
         */
        this.sharedStringsObj = {} // 共享字符串去重对象清空
        /**
         * XML Template
         * @type {Object}
         */
        this.objTp = snXml
        this._styleCache = []
    }

    /**
     * Retrieving .rels relationships by relational taget path
     * @param {string} targetPath - Target Path
     * @param {boolean} [create=false] - Whether to create when it does not exist
     * @returns {Object}
     */
    getRelsByTarget(targetPath, create = false) {
        // 1. 规范化路径，去除所有 "../"
        const normalized = targetPath.split('/').filter(segment => segment !== '..').join('/');
        // 2. 提取文件名和目录
        const fileName = normalized.split('/').pop();           // e.g. "drawing1.xml"
        const dir = normalized.replace(`/${fileName}`, '');     // e.g. "drawings"
        // 3. 构造 rels 文件键，避免重复添加 "xl/"
        const basePath = normalized.startsWith('xl/') ? '' : 'xl/';
        const relsKey = `${basePath}${dir}/_rels/${fileName}.rels`;
        let resource = this.obj[relsKey]
        // 如果不存在依赖，是否创建一个
        if (!resource && create) {
            resource = this.obj[relsKey] = {
                "?xml": {
                    "_$version": "1.0",
                    "_$encoding": "UTF-8",
                    "_$standalone": "yes"
                },
                "Relationships": {
                    "Relationship": [],
                    "_$xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
                }
            }
        }
        return resource?.Relationships?.Relationship;
    }

    /**
     * Reliance RId from a sheet to get the target path
     * @param {string} sheetRId - Sheet
     * @param {string} fileRId - Reliance file RId
     * @returns {string|null}
     */
    getWsRelsFileTarget(sheetRId, fileRId) {
        const sheetRel = this.relationship.find(r => r._$Id === sheetRId);
        if (!sheetRel) return null;

        const targetPath = sheetRel._$Target; // e.g. "worksheets/sheet1.xml"
        const relations = this.getRelsByTarget(targetPath);
        if (!relations) return null;

        const fileRel = relations.find(r => r._$Id === fileRId);
        if (!fileRel) return null;

        let fileKey = `xl/${fileRel._$Target.replace(/\.\.\//g, '')}`;

        return fileKey
    }

    /**
     * Get theme color by theme index
     * @param {number} themeIndex - Theme index
     * @param {number} [tint=0] - Tint value
     * @returns {string}
     */
    getThemeColor(themeIndex, tint = 0) {
        return _getThemeColor(themeIndex, tint, this.SN);
    }

    /**
     * Get indexed color by index
     * @param {number} index - Indexed color index
     * @returns {string}
     */
    getIndexedColor(index) {
        return _getColorByIndex(index);
    }

    /**
     * Shared String Table
     * @type {Object}     */
    /**
     * Shared String Table
     * @type {Object}     */

    get sst() {
        const sst = this.obj['xl/sharedStrings.xml']
        if (!sst) { // 不存在则创建
            this.obj["xl/sharedStrings.xml"] = {
                "?xml": {
                    "_$version": "1.0",
                    "_$encoding": "UTF-8",
                    "_$standalone": "yes"
                },
                "sst": {
                    "si": [],
                    "_$xmlns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "_$count": "0",
                    "_$uniqueCount": "0"
                }
            }
            this.override.push({
                "_$PartName": "/xl/sharedStrings.xml",
                "_$ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
            })
            this.relationship.push({
                "_$Id": this._xGetNextRId(),
                "_$Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                "_$Target": "sharedStrings.xml"
            })
        }
        return this.obj['xl/sharedStrings.xml'].sst
    }

    /**
     * Stylesheet
     * @type {Object}     */
    /**
     * Stylesheet
     * @type {Object}     */

    get styleSheet() {
        return this.obj['xl/styles.xml'].styleSheet
    }

    /**
     * Reset Stylesheet
     * @returns {void}
     */
    resetStyle() {
        this.obj['xl/styles.xml'] = JSON.parse(JSON.stringify(this.objTp['xl/styles.xml']))
    }

    /**
     * Overwrite Relationship
     * @type {Array}     */
    /**
     * Overwrite Relationship
     * @type {Array}     */

    get override() {
        return this.obj["[Content_Types].xml"].Types.Override
    }

    /**
     * Relationship collection
     * @type {Array}     */
    /**
     * Relationship collection
     * @type {Array}     */

    get relationship() {
        return this.obj["xl/_rels/workbook.xml.rels"].Relationships.Relationship
    }

    /**
     * Current File Theme Colour
     * @type {Object}     */
    /**
     * Current File Theme Colour
     * @type {Object}     */

    get clrScheme() { // 当前文件主题
        return this.obj?.["xl/theme/theme1.xml"]?.["a:theme"]?.["a:themeElements"]?.["a:clrScheme"]
    }
    /**
     * All Sheet Information
     * @type {Array}     */
    /**
     * All Sheet Information
     * @type {Array}     */

    get sheets() { // 获取所有xml表格信息
        return this.obj["xl/workbook.xml"].workbook.sheets.sheet
    }
    /**
     * Theme Template Colouring
     * @type {Object}     */
    /**
     * Theme Template Colouring
     * @type {Object}     */

    get clrSchemeTp() { // 主题模板
        return this.objTp["xl/theme/theme1.xml"]["a:theme"]["a:themeElements"]["a:clrScheme"]
    }
    /**
     * Add Sheet
     * @param {string} sheetName - Sheet Name
     * @param {Object} ops - Initializing parameters
     * @returns {void}
     */
    addSheet(sheetName, ops) {
        // sheetId生成
        const sheetId = Math.max(...this.sheets.map(st => Number(st._$sheetId))) + 1;
        // 工作表文件名生成
        const sheetKey = this._xGetNextSheetFileKey();
        // 新rId
        const rId = this._xGetNextRId();
        // 添加工作簿
        const ctMeta = { "_$name": sheetName, "_$sheetId": sheetId, "_$r:id": rId }
        if (ops.hidden) ctMeta._$state = "hidden"
        this.sheets.push(ctMeta)
        // 添加关系文件
        this.relationship.push({
            "_$Id": rId,
            "_$Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            "_$Target": sheetKey.replace("xl/", "")
        })
        // 添加文件类型声明
        this.override.push({
            "_$PartName": "/" + sheetKey,
            "_$ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
        })
        // 添加最终key
        return this.obj[sheetKey] = JSON.parse(JSON.stringify(this.objTp['xl/worksheets/sheet1.xml']));
    }
    // 获取单元格样式 - 极致性能优化版本
    _xGetStyle(styleIndex) {
        const styleObj = this.obj['xl/styles.xml'].styleSheet;
        const xf = styleObj.cellXfs.xf[styleIndex];

        // 预先获取所有需要的对象引用，减少属性访问
        const fontId = xf['_$fontId'];
        const borderId = xf['_$borderId'];
        const fillId = xf['_$fillId'];
        const numFmtId = xf['_$numFmtId'];

        const font = styleObj.fonts?.font?.[fontId] || {};
        const alignment = xf.alignment;
        const border = styleObj.borders?.border?.[borderId] || {};
        const fill = styleObj.fills?.fill?.[fillId] || {};

        // 预计算填充类型
        const patternFill = fill.patternFill;
        const gradientFill = fill.gradientFill;
        const fType = patternFill ? 'pattern' : gradientFill ? 'gradient' : null;

        const result = {};

        result.y = { font, alignment, border, fill }

        // 字体样式处理 - 只添加存在的属性
        const fontStyle = {};
        let hasFontStyle = false;

        if (font.b === "") { fontStyle.bold = true; hasFontStyle = true; }
        if (font.i === "") { fontStyle.italic = true; hasFontStyle = true; }
        if (font.strike === "") { fontStyle.strike = true; hasFontStyle = true; }
        if (font.outline === "") { fontStyle.outline = true; hasFontStyle = true; }

        const fontName = font.name?._$val;
        if (fontName) { fontStyle.name = fontName; hasFontStyle = true; }

        const fontCharset = font.charset?._$val;
        if (fontCharset) { fontStyle.charset = fontCharset; hasFontStyle = true; }

        const fontSize = font.sz?._$val;
        if (fontSize) {
            const parsedFontSize = Number(fontSize);
            fontStyle.size = Number.isFinite(parsedFontSize) ? parsedFontSize : fontSize;
            hasFontStyle = true;
        }

        const fontColorNode = font.color;
        const fontColor = getColor(fontColorNode, this.SN);
        if (fontColor) { fontStyle.color = fontColor; hasFontStyle = true; }
        if (fontColorNode?._$theme !== undefined) {
            fontStyle.colorTheme = fontColorNode._$theme;
            if (fontColorNode._$tint !== undefined) fontStyle.colorTint = fontColorNode._$tint;
            hasFontStyle = true;
        } else if (fontColorNode?._$indexed !== undefined) {
            fontStyle.colorIndexed = fontColorNode._$indexed;
            hasFontStyle = true;
        } else if (fontColorNode?._$auto !== undefined) {
            fontStyle.colorAuto = fontColorNode._$auto;
            hasFontStyle = true;
        }

        // 下划线处理
        if (font.u !== undefined) {
            fontStyle.underline = typeof font.u === 'object' ? font.u._$val : 'single';
            hasFontStyle = true;
        }

        const vertAlign = font.vertAlign?._$val;
        if (vertAlign) { fontStyle.vertAlign = vertAlign; hasFontStyle = true; }

        if (hasFontStyle) result.font = fontStyle;

        // 数字格式处理 - 优化查找逻辑
        if (numFmtId) {
            let numFmt = BUILTIN_NUMFMTS[numFmtId];
            if (!numFmt && styleObj.numFmts?.numFmt) {
                // 只在内置格式不存在时才查找自定义格式
                const customFmts = styleObj.numFmts.numFmt;
                for (let i = 0, len = customFmts.length; i < len; i++) {
                    if (customFmts[i]['_$numFmtId'] == numFmtId) {
                        numFmt = customFmts[i]['_$formatCode'];
                        break;
                    }
                }
            }
            if (numFmt) {
                // 优化字符串替换 - 只在需要时执行
                result.numFmt = numFmt.indexOf('\\') !== -1 ? numFmt.replaceAll("\\", "") : numFmt;
            }
        }

        // 对齐样式处理
        if (alignment) {
            const alignStyle = {};
            let hasAlignStyle = false;

            const horizontal = alignment._$horizontal;
            if (horizontal) { alignStyle.horizontal = horizontal; hasAlignStyle = true; }

            const vertical = alignment._$vertical;
            if (vertical) { alignStyle.vertical = vertical; hasAlignStyle = true; }

            if (alignment._$wrapText === '1') { alignStyle.wrapText = true; hasAlignStyle = true; }

            if (alignment._$shrinkToFit === '1') { alignStyle.shrinkToFit = true; hasAlignStyle = true; }

            const indent = alignment._$indent;
            if (indent) {
                const parsedIndent = Number(indent);
                alignStyle.indent = Number.isFinite(parsedIndent) ? parsedIndent : indent;
                hasAlignStyle = true;
            }

            const textRotation = alignment._$textRotation;
            if (textRotation) { alignStyle.textRotation = parseInt(textRotation); hasAlignStyle = true; }

            if (hasAlignStyle) result.alignment = alignStyle;
        }

        // 边框样式处理
        const borderStyle = {};
        let hasBorderStyle = false;

        const borderSides = ['bottom', 'top', 'left', 'right'];
        for (let i = 0; i < 4; i++) {
            const side = borderSides[i];
            const borderSide = border[side];
            if (borderSide) {
                const color = getColor(borderSide.color, this.SN);
                const style = borderSide._$style;
                if (color || style) {
                    borderStyle[side] = {};
                    if (color) borderStyle[side].color = color;
                    if (style) borderStyle[side].style = style;
                    hasBorderStyle = true;
                }
            }
        }

        const diagonalSide = border.diagonal;
        const diagonalDown = border._$diagonalDown === '1';
        const diagonalUp = border._$diagonalUp === '1';
        const hasDiagonalFlag = diagonalDown || diagonalUp;
        if (diagonalSide && (diagonalSide._$style || diagonalSide.color || hasDiagonalFlag)) {
            const diagonalObj = {};
            const diagonalColor = getColor(diagonalSide.color, this.SN);
            if (diagonalColor) diagonalObj.color = diagonalColor;
            if (diagonalSide._$style) diagonalObj.style = diagonalSide._$style;
            borderStyle.diagonal = diagonalObj;
            if (diagonalDown || !hasDiagonalFlag) borderStyle.diagonalDown = true;
            if (diagonalUp) borderStyle.diagonalUp = true;
            hasBorderStyle = true;
        }

        if (hasBorderStyle) result.border = borderStyle;

        // 填充样式处理
        if (fType) {
            const fillStyle = { type: fType };
            let hasVisibleFill = false;

            if (fType === 'pattern') {
                const pattern = patternFill._$patternType;
                if (pattern) {
                    fillStyle.pattern = pattern;
                    if (pattern !== 'none') hasVisibleFill = true;
                }

                const fgColor = getColor(patternFill.fgColor, this.SN);
                if (fgColor) {
                    fillStyle.fgColor = fgColor;
                    hasVisibleFill = true;
                }

                const bgColor = getColor(patternFill.bgColor, this.SN);
                if (bgColor) {
                    fillStyle.bgColor = bgColor;
                    hasVisibleFill = true;
                }
            } else if (fType === 'gradient') {
                hasVisibleFill = true;
                const gradientType = gradientFill._$type;
                if (gradientType) fillStyle.gradientType = gradientType;

                const degree = gradientFill._$degree;
                if (degree) {
                    const parsedDegree = Number(degree);
                    fillStyle.degree = Number.isFinite(parsedDegree) ? parsedDegree : degree;
                }

                const left = gradientFill._$left;
                if (left) {
                    const parsedLeft = Number(left);
                    fillStyle.left = Number.isFinite(parsedLeft) ? parsedLeft : left;
                }

                const right = gradientFill._$right;
                if (right) {
                    const parsedRight = Number(right);
                    fillStyle.right = Number.isFinite(parsedRight) ? parsedRight : right;
                }

                const top = gradientFill._$top;
                if (top) {
                    const parsedTop = Number(top);
                    fillStyle.top = Number.isFinite(parsedTop) ? parsedTop : top;
                }

                const bottom = gradientFill._$bottom;
                if (bottom) {
                    const parsedBottom = Number(bottom);
                    fillStyle.bottom = Number.isFinite(parsedBottom) ? parsedBottom : bottom;
                }

                const stops = gradientFill.stop;
                if (stops) {
                    fillStyle.stops = [];
                    for (let i = 0, len = stops.length; i < len; i++) {
                        const stop = stops[i];
                        const stopColor = getColor(stop.color, this.SN);
                        if (stopColor || stop._$position) {
                            const stopObj = {};
                            if (stop._$position) {
                                const parsedPosition = Number(stop._$position);
                                stopObj.position = Number.isFinite(parsedPosition) ? parsedPosition : stop._$position;
                            }
                            if (stopColor) stopObj.color = stopColor;
                            fillStyle.stops.push(stopObj);
                        }
                    }
                }
            }

            if (hasVisibleFill) {
                result.fill = fillStyle;
            }
        }

        const protection = xf.protection;
        if (protection) {
            const protectionObj = {};
            if (protection._$locked !== undefined) {
                protectionObj.locked = protection._$locked !== '0';
            }
            if (protection._$hidden !== undefined) {
                protectionObj.hidden = protection._$hidden === '1';
            }
            if (Object.keys(protectionObj).length > 0) {
                result.protection = protectionObj;
            }
        }

        return this._styleCache[styleIndex] = result;
    }

    // 获取差异化样式（用于条件格式）
    _xGetDxfStyle(dxfId) {
        const styleObj = this.obj['xl/styles.xml'].styleSheet;
        const dxf = styleObj.dxfs?.dxf;
        if (!dxf) return null;

        const dxfArr = Array.isArray(dxf) ? dxf : [dxf];
        const item = dxfArr[dxfId];
        if (!item) return null;

        const result = {};
        const isEnabledFlag = (node) => {
            if (node === undefined || node === null) return false;
            if (node === '') return true;
            if (typeof node === 'object') {
                const val = node._$val;
                if (val === undefined || val === null || val === '') return true;
                const lowered = String(val).toLowerCase();
                return lowered !== '0' && lowered !== 'false' && lowered !== 'none' && lowered !== 'off';
            }
            const lowered = String(node).toLowerCase();
            return lowered !== '0' && lowered !== 'false' && lowered !== 'none' && lowered !== 'off';
        };

        // 字体样式
        if (item.font) {
            const font = item.font;
            const fontStyle = {};
            let hasFontStyle = false;

            if (isEnabledFlag(font.b)) { fontStyle.bold = true; hasFontStyle = true; }
            if (isEnabledFlag(font.i)) { fontStyle.italic = true; hasFontStyle = true; }
            if (isEnabledFlag(font.strike)) { fontStyle.strike = true; hasFontStyle = true; }

            const fontColor = getColor(font.color, this.SN);
            if (fontColor) { fontStyle.color = fontColor; hasFontStyle = true; }

            if (font.u !== undefined && isEnabledFlag(font.u)) {
                fontStyle.underline = typeof font.u === 'object' ? font.u._$val : 'single';
                hasFontStyle = true;
            }

            if (hasFontStyle) result.font = fontStyle;
        }

        // 填充样式
        if (item.fill) {
            const patternFill = item.fill.patternFill;
            if (patternFill) {
                const fillStyle = { type: 'pattern' };
                const pattern = patternFill._$patternType;
                if (pattern) fillStyle.pattern = pattern;

                // dxf中bgColor是背景色，但渲染时用fgColor
                const bgColor = getColor(patternFill.bgColor, this.SN);
                if (bgColor) fillStyle.fgColor = bgColor;

                const fgColor = getColor(patternFill.fgColor, this.SN);
                if (fgColor && !bgColor) fillStyle.fgColor = fgColor;

                result.fill = fillStyle;
            }
        }

        // 边框样式
        if (item.border) {
            const border = item.border;
            const borderStyle = {};
            let hasBorderStyle = false;

            const borderSides = ['bottom', 'top', 'left', 'right'];
            for (const side of borderSides) {
                const borderSide = border[side];
                if (borderSide) {
                    const color = getColor(borderSide.color, this.SN);
                    const style = borderSide._$style;
                    if (color || style) {
                        borderStyle[side] = {};
                        if (color) borderStyle[side].color = color;
                        if (style) borderStyle[side].style = style;
                        hasBorderStyle = true;
                    }
                }
            }

            const diagonalSide = border.diagonal;
            const diagonalDown = border._$diagonalDown === '1';
            const diagonalUp = border._$diagonalUp === '1';
            const hasDiagonalFlag = diagonalDown || diagonalUp;
            if (diagonalSide && (diagonalSide._$style || diagonalSide.color || hasDiagonalFlag)) {
                const diagonalObj = {};
                const diagonalColor = getColor(diagonalSide.color, this.SN);
                if (diagonalColor) diagonalObj.color = diagonalColor;
                if (diagonalSide._$style) diagonalObj.style = diagonalSide._$style;
                borderStyle.diagonal = diagonalObj;
                if (diagonalDown || !hasDiagonalFlag) borderStyle.diagonalDown = true;
                if (diagonalUp) borderStyle.diagonalUp = true;
                hasBorderStyle = true;
            }

            if (hasBorderStyle) result.border = borderStyle;
        }

        return result;
    }

    // 获取一个表的xmlObj通过rId
    _xGetSheetXmlObj(rId) {
        const findObj = this.relationship.find(item => item._$Id == rId);
        if (!findObj) return null
        return this.obj[`xl/${findObj._$Target.replace("/xl/", "")}`].worksheet;
    }
    // 获取工作表文件名通过rId
    _getSheetFileName(rId) {
        const findObj = this.relationship.find(item => item._$Id == rId);
        if (!findObj) return null;
        // 返回类似 "sheet1.xml" 的文件名
        return findObj._$Target.split("/").pop();
    }
    // 解析单个 run 的字体属性
    _parseRunFont(rPr) {
        if (!rPr) return null;
        const font = {};
        let has = false;
        if (rPr.b === "") { font.bold = true; has = true; }
        if (rPr.i === "") { font.italic = true; has = true; }
        if (rPr.strike === "") { font.strike = true; has = true; }
        if (rPr.u !== undefined) {
            font.underline = typeof rPr.u === 'object' ? rPr.u._$val : 'single';
            has = true;
        }
        const sz = rPr.sz?._$val;
        if (sz) {
            const parsedSize = Number(sz);
            font.size = Number.isFinite(parsedSize) ? parsedSize : sz;
            has = true;
        }
        const color = getColor(rPr.color, this.SN);
        if (color) { font.color = color; has = true; }
        const name = rPr.rFont?._$val;
        if (name) { font.name = name; has = true; }
        const charset = rPr.charset?._$val;
        if (charset) { font.charset = charset; has = true; }
        const vertAlign = rPr.vertAlign?._$val;
        if (vertAlign) { font.vertAlign = vertAlign; has = true; }
        return has ? font : null;
    }
    // 获取共享字符
    _xGetSharedStrings(v) {
        const val = this.obj["xl/sharedStrings.xml"]?.sst?.si?.[v];
        let str = "";
        let richText = null;
        if (typeof val.t === 'string') {
            str = val.t;
        } else if (val.r) {
            const rArr = Array.isArray(val.r) ? val.r : [val.r];
            const runs = [];
            let hasFormat = false;
            for (const item of rArr) {
                const text = typeof item.t === 'string' ? item.t : (item.t?.['#text'] ?? "");
                const font = this._parseRunFont(item.rPr);
                if (font) hasFormat = true;
                runs.push({ text, font: font || {} });
            }
            str = runs.map(r => r.text).join("");
            if (hasFormat) richText = runs;
        } else {
            str = val.t?.['#text'] ?? val?.t ?? "";
        }
        // 判断并输出调试信息
        if (typeof str !== 'string') return { text: "", richText: null };

        str = str.replaceAll("&#10;", "\n");
        if (richText) {
            for (const run of richText) {
                run.text = run.text.replaceAll("&#10;", "\n");
            }
        }
        return { text: str, richText };
    }
    // 获取超链接target信息
    /**
     * Parsing Hyperlinks Based on RId
     * @param {string} rId - Relationship Id
     * @returns {Object|null}
     */
    resolveHyperlinkById(rId) {
        const findObj = this.relationship.find(item => item._$Id == rId); // 找到文件名
        if (!findObj) return null;
        const rel = this.obj[`xl/worksheets/_rels/${findObj._$Target.split("/").pop()}.rels`]
        if (!rel?.Relationships?.Relationship) return null
        const obj = {}
        const rels = Array.isArray(rel.Relationships.Relationship) ? rel.Relationships.Relationship : [rel.Relationships.Relationship];
        rels.forEach(t => {
            obj[t._$Id] = t
        })
        return obj
    }
    // 获取下一个rId
    _xGetNextRId() {
        const rIds = this.relationship.map(item => parseInt(item._$Id.match(/\d+/)[0]))
        return `rId${Math.max(...rIds) + 1}`;
    }
    // 获取下一个工作表key
    _xGetNextSheetFileKey() {
        const keys = Object.keys(this.obj);
        let maxNumber = 0;
        keys.forEach(key => {
            const match = key.match(/^xl\/worksheets\/sheet(\d+)\.xml$/);
            if (match) {
                const num = parseInt(match[1], 10);
                if (num > maxNumber) maxNumber = num;
            }
        });
        const nextNumber = maxNumber + 1; // 最大数字加 1
        return `xl/worksheets/sheet${nextNumber}.xml`; // 返回新键
    }
}
