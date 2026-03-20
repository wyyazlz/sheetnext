export const DEFAULT_CELL_FONT_NAME = '宋体';
export const DEFAULT_CELL_FONT_SIZE = 11;
export const DEFAULT_CELL_FONT_CHARSET = 134;

function _resolveSheetNext(source) {
    return source?.SN || source?._SN || source || null;
}

function _getWorkbookDefaultFontNode(source) {
    const SN = _resolveSheetNext(source);
    const fontNode = SN?._Xml?.styleSheet?.fonts?.font ?? SN?.Xml?.styleSheet?.fonts?.font;
    if (Array.isArray(fontNode)) return fontNode[0] || null;
    return fontNode || null;
}

/** @param {any} source @returns {{name:string,size:number,charset:number}} */
export function getDefaultCellFont(source) {
    const fontNode = _getWorkbookDefaultFontNode(source);
    const size = Number(fontNode?.sz?._$val);
    const charsetRaw = fontNode?.charset?._$val;
    const charset = charsetRaw === undefined || charsetRaw === null || charsetRaw === ''
        ? DEFAULT_CELL_FONT_CHARSET
        : Number(charsetRaw);
    return {
        name: fontNode?.name?._$val || DEFAULT_CELL_FONT_NAME,
        size: size > 0 ? size : DEFAULT_CELL_FONT_SIZE,
        charset: Number.isFinite(charset) ? charset : DEFAULT_CELL_FONT_CHARSET
    };
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
