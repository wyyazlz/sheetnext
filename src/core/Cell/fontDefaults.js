export const DEFAULT_CELL_FONT_NAME = '宋体';
export const DEFAULT_CELL_FONT_SIZE = 11;
export const DEFAULT_CELL_FONT_CHARSET = 134;

const EN_DEFAULT_CELL_FONT_NAME = 'Calibri';
const EN_DEFAULT_CELL_FONT_CHARSET = 0;

function _resolveSheetNext(source) {
    return source?.SN || source?._SN || source || null;
}

function _getLocaleDefaultCellFont(source) {
    const locale = String(_resolveSheetNext(source)?.locale || '');
    if (/^en\b/i.test(locale)) {
        return {
            name: EN_DEFAULT_CELL_FONT_NAME,
            size: DEFAULT_CELL_FONT_SIZE,
            charset: EN_DEFAULT_CELL_FONT_CHARSET
        };
    }
    return {
        name: DEFAULT_CELL_FONT_NAME,
        size: DEFAULT_CELL_FONT_SIZE,
        charset: DEFAULT_CELL_FONT_CHARSET
    };
}

function _getWorkbookDefaultFontNode(source) {
    const SN = _resolveSheetNext(source);
    const fontNode = SN?.Xml?.styleSheet?.fonts?.font ?? SN?._Xml?.styleSheet?.fonts?.font;
    if (Array.isArray(fontNode)) return fontNode[0] || null;
    return fontNode || null;
}

function _getWritableStyleSheet(source) {
    const SN = _resolveSheetNext(source);
    return source?.styleSheet
        || source?.obj?.['xl/styles.xml']?.styleSheet
        || SN?.Xml?.styleSheet
        || SN?._Xml?.styleSheet
        || null;
}

function _setDefaultFontNode(fontNode, preset, size) {
    if (!fontNode || !preset) return;
    fontNode.sz = { _$val: String(size) };
    fontNode.name = { _$val: preset.name };
    fontNode.charset = { _$val: String(preset.charset) };
    fontNode.scheme = { _$val: 'minor' };
}

/** @param {any} source @returns {{name:string,size:number,charset:number}} */
export function getDefaultCellFont(source) {
    const preset = _getLocaleDefaultCellFont(source);
    const fontNode = _getWorkbookDefaultFontNode(source);
    const size = Number(fontNode?.sz?._$val);
    const charsetRaw = fontNode?.charset?._$val;
    const charset = charsetRaw === undefined || charsetRaw === null || charsetRaw === ''
        ? preset.charset
        : Number(charsetRaw);
    return {
        name: fontNode?.name?._$val || preset.name,
        size: size > 0 ? size : preset.size,
        charset: Number.isFinite(charset) ? charset : preset.charset
    };
}

/** @param {any} source */
export function applyLocaleDefaultCellFont(source) {
    const styleSheet = _getWritableStyleSheet(source);
    if (!styleSheet) return;
    if (!styleSheet.fonts || typeof styleSheet.fonts !== 'object') {
        styleSheet.fonts = { font: [], '_$count': '0' };
    }

    const preset = _getLocaleDefaultCellFont(source);
    const fontNodes = Array.isArray(styleSheet.fonts.font)
        ? styleSheet.fonts.font
        : (styleSheet.fonts.font ? [styleSheet.fonts.font] : []);

    while (fontNodes.length < 2) fontNodes.push({});
    _setDefaultFontNode(fontNodes[0], preset, preset.size);
    _setDefaultFontNode(fontNodes[1], preset, 9);
    if (!fontNodes[0].color) fontNodes[0].color = { _$theme: '1' };

    styleSheet.fonts.font = fontNodes;
    styleSheet.fonts['_$count'] = String(fontNodes.length);
    if (!styleSheet.fonts['_$x14ac:knownFonts']) styleSheet.fonts['_$x14ac:knownFonts'] = '1';
}

/** @param {any} source @returns {string} */
export function getDefaultCellFontName(source) {
    return getDefaultCellFont(source).name;
}

/** @param {any} source @returns {number} */
export function getDefaultCellFontSize(source) {
    return getDefaultCellFont(source).size;
}

/** @param {Object} font @param {any} source @returns {string} */
export function getFontName(font, source) {
    return font?.name || getDefaultCellFont(source).name;
}

/** @param {Object} font @param {any} source @returns {number} */
export function getFontSize(font, source) {
    const size = Number(font?.size);
    return size > 0 ? size : getDefaultCellFont(source).size;
}

/** @param {Object} font @param {any} source @returns {number} */
export function getFontSizePx(font, source) {
    return getFontSize(font, source) * 1.333;
}

/** @param {Object} font @param {any} source @param {number} [sizePx] @returns {string} */
export function getCanvasFont(font, source, sizePx) {
    const resolvedSize = sizePx ?? getFontSizePx(font, source);
    const bold = font?.bold ? 'bold' : '';
    const italic = font?.italic ? 'italic' : '';
    return `${bold ? 'bold ' : ''}${italic ? 'italic ' : ''}${resolvedSize}px ${getFontName(font, source)}`.trim();
}
