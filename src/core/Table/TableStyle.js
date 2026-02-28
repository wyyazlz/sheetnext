import { getThemeColor } from '../Xml/helpers.js';
import { TABLE_STYLE_TOKENS } from './TableStyleTokens.js';

/**
 * Fallback style palette (used when theme/token resolution is not available).
 * Naming follows: TableStyle + Light/Medium/Dark + number.
 */
const THEME_COLORS = {
    blue: { light: '#DCE6F1', medium: '#4F81BD', dark: '#1F497D' },
    orange: { light: '#FDE9D9', medium: '#F79646', dark: '#E26B0A' },
    gray: { light: '#D9D9D9', medium: '#808080', dark: '#404040' },
    yellow: { light: '#FFF2CC', medium: '#FFCC00', dark: '#BF9000' },
    lightBlue: { light: '#DAEEF3', medium: '#4BACC6', dark: '#1D6A7C' },
    green: { light: '#EBF1DE', medium: '#9BBB59', dark: '#4F6228' },
    purple: { light: '#E5DFEC', medium: '#8064A2', dark: '#4C3B5F' },
    red: { light: '#F2DCDB', medium: '#C0504D', dark: '#952D2A' }
};

function createStyle(name, colors, type) {
    const { light, medium, dark } = colors;

    if (type === 'light') {
        return {
            name,
            headerBg: null,
            headerFg: dark,
            headerBorder: medium,
            rowStripe1: null,
            rowStripe2: light,
            border: '#D9D9D9',
            firstColumnBg: null,
            lastColumnBg: null,
            totalsBg: null
        };
    }

    if (type === 'medium') {
        return {
            name,
            headerBg: medium,
            headerFg: '#FFFFFF',
            headerBorder: medium,
            rowStripe1: '#FFFFFF',
            rowStripe2: light,
            border: medium,
            firstColumnBg: light,
            lastColumnBg: light,
            totalsBg: light
        };
    }

    return {
        name,
        headerBg: dark,
        headerFg: '#FFFFFF',
        headerBorder: dark,
        rowStripe1: medium,
        rowStripe2: light,
        border: dark,
        firstColumnBg: medium,
        lastColumnBg: medium,
        totalsBg: medium
    };
}

/**
 * Fallback table styles.
 */
export const TABLE_STYLES = {
    // Light (1-21)
    TableStyleLight1: createStyle('TableStyleLight1', THEME_COLORS.gray, 'light'),
    TableStyleLight2: createStyle('TableStyleLight2', THEME_COLORS.blue, 'light'),
    TableStyleLight3: createStyle('TableStyleLight3', THEME_COLORS.orange, 'light'),
    TableStyleLight4: createStyle('TableStyleLight4', THEME_COLORS.gray, 'light'),
    TableStyleLight5: createStyle('TableStyleLight5', THEME_COLORS.yellow, 'light'),
    TableStyleLight6: createStyle('TableStyleLight6', THEME_COLORS.lightBlue, 'light'),
    TableStyleLight7: createStyle('TableStyleLight7', THEME_COLORS.green, 'light'),
    TableStyleLight8: createStyle('TableStyleLight8', THEME_COLORS.blue, 'light'),
    TableStyleLight9: createStyle('TableStyleLight9', THEME_COLORS.orange, 'light'),
    TableStyleLight10: createStyle('TableStyleLight10', THEME_COLORS.gray, 'light'),
    TableStyleLight11: createStyle('TableStyleLight11', THEME_COLORS.blue, 'light'),
    TableStyleLight12: createStyle('TableStyleLight12', THEME_COLORS.orange, 'light'),
    TableStyleLight13: createStyle('TableStyleLight13', THEME_COLORS.yellow, 'light'),
    TableStyleLight14: createStyle('TableStyleLight14', THEME_COLORS.lightBlue, 'light'),
    TableStyleLight15: createStyle('TableStyleLight15', THEME_COLORS.green, 'light'),
    TableStyleLight16: createStyle('TableStyleLight16', THEME_COLORS.purple, 'light'),
    TableStyleLight17: createStyle('TableStyleLight17', THEME_COLORS.red, 'light'),
    TableStyleLight18: createStyle('TableStyleLight18', THEME_COLORS.blue, 'light'),
    TableStyleLight19: createStyle('TableStyleLight19', THEME_COLORS.orange, 'light'),
    TableStyleLight20: createStyle('TableStyleLight20', THEME_COLORS.gray, 'light'),
    TableStyleLight21: createStyle('TableStyleLight21', THEME_COLORS.blue, 'light'),

    // Medium (1-28)
    TableStyleMedium1: createStyle('TableStyleMedium1', THEME_COLORS.gray, 'medium'),
    TableStyleMedium2: createStyle('TableStyleMedium2', THEME_COLORS.blue, 'medium'),
    TableStyleMedium3: createStyle('TableStyleMedium3', THEME_COLORS.orange, 'medium'),
    TableStyleMedium4: createStyle('TableStyleMedium4', THEME_COLORS.gray, 'medium'),
    TableStyleMedium5: createStyle('TableStyleMedium5', THEME_COLORS.yellow, 'medium'),
    TableStyleMedium6: createStyle('TableStyleMedium6', THEME_COLORS.lightBlue, 'medium'),
    TableStyleMedium7: createStyle('TableStyleMedium7', THEME_COLORS.green, 'medium'),
    TableStyleMedium8: createStyle('TableStyleMedium8', THEME_COLORS.blue, 'medium'),
    TableStyleMedium9: createStyle('TableStyleMedium9', THEME_COLORS.orange, 'medium'),
    TableStyleMedium10: createStyle('TableStyleMedium10', THEME_COLORS.gray, 'medium'),
    TableStyleMedium11: createStyle('TableStyleMedium11', THEME_COLORS.blue, 'medium'),
    TableStyleMedium12: createStyle('TableStyleMedium12', THEME_COLORS.orange, 'medium'),
    TableStyleMedium13: createStyle('TableStyleMedium13', THEME_COLORS.yellow, 'medium'),
    TableStyleMedium14: createStyle('TableStyleMedium14', THEME_COLORS.lightBlue, 'medium'),
    TableStyleMedium15: createStyle('TableStyleMedium15', THEME_COLORS.green, 'medium'),
    TableStyleMedium16: createStyle('TableStyleMedium16', THEME_COLORS.purple, 'medium'),
    TableStyleMedium17: createStyle('TableStyleMedium17', THEME_COLORS.red, 'medium'),
    TableStyleMedium18: createStyle('TableStyleMedium18', THEME_COLORS.blue, 'medium'),
    TableStyleMedium19: createStyle('TableStyleMedium19', THEME_COLORS.orange, 'medium'),
    TableStyleMedium20: createStyle('TableStyleMedium20', THEME_COLORS.gray, 'medium'),
    TableStyleMedium21: createStyle('TableStyleMedium21', THEME_COLORS.blue, 'medium'),
    TableStyleMedium22: createStyle('TableStyleMedium22', THEME_COLORS.orange, 'medium'),
    TableStyleMedium23: createStyle('TableStyleMedium23', THEME_COLORS.gray, 'medium'),
    TableStyleMedium24: createStyle('TableStyleMedium24', THEME_COLORS.yellow, 'medium'),
    TableStyleMedium25: createStyle('TableStyleMedium25', THEME_COLORS.lightBlue, 'medium'),
    TableStyleMedium26: createStyle('TableStyleMedium26', THEME_COLORS.green, 'medium'),
    TableStyleMedium27: createStyle('TableStyleMedium27', THEME_COLORS.purple, 'medium'),
    TableStyleMedium28: createStyle('TableStyleMedium28', THEME_COLORS.red, 'medium'),

    // Dark (1-11)
    TableStyleDark1: createStyle('TableStyleDark1', THEME_COLORS.gray, 'dark'),
    TableStyleDark2: createStyle('TableStyleDark2', THEME_COLORS.blue, 'dark'),
    TableStyleDark3: createStyle('TableStyleDark3', THEME_COLORS.orange, 'dark'),
    TableStyleDark4: createStyle('TableStyleDark4', THEME_COLORS.gray, 'dark'),
    TableStyleDark5: createStyle('TableStyleDark5', THEME_COLORS.yellow, 'dark'),
    TableStyleDark6: createStyle('TableStyleDark6', THEME_COLORS.lightBlue, 'dark'),
    TableStyleDark7: createStyle('TableStyleDark7', THEME_COLORS.green, 'dark'),
    TableStyleDark8: createStyle('TableStyleDark8', THEME_COLORS.blue, 'dark'),
    TableStyleDark9: createStyle('TableStyleDark9', THEME_COLORS.orange, 'dark'),
    TableStyleDark10: createStyle('TableStyleDark10', THEME_COLORS.purple, 'dark'),
    TableStyleDark11: createStyle('TableStyleDark11', THEME_COLORS.red, 'dark')
};

const STYLE_ELEMENT = {
    wholeTable: '0',
    headerRow: '1',
    totalRow: '2',
    firstColumn: '3',
    lastColumn: '4',
    rowStripe1: '5',
    rowStripe2: '6',
    columnStripe1: '7',
    columnStripe2: '8'
};

function pickDefined(...values) {
    for (const value of values) {
        if (value !== undefined) return value;
    }
    return undefined;
}

function resolveTokenColor(token, SN) {
    if (Array.isArray(token) && token.length >= 2) {
        if (!SN?.Xml) return undefined;
        const [themeIndex, tint] = token;
        try {
            return getThemeColor(themeIndex, tint, SN);
        } catch (_e) {
            return undefined;
        }
    }

    if (typeof token === 'string') return token;
    return undefined;
}

function resolveElement(tokens, key, SN) {
    const token = tokens?.[key];
    if (!token) {
        return {
            hasFormat: false,
            bgColor: undefined,
            fgColor: undefined,
            isBold: undefined,
            stripeSize: 1
        };
    }

    const hasFormat = token.h === 1;
    if (!hasFormat) {
        return {
            hasFormat: false,
            bgColor: undefined,
            fgColor: undefined,
            isBold: undefined,
            stripeSize: token.s || 1
        };
    }

    const hasFill = Object.prototype.hasOwnProperty.call(token, 'fb');
    const hasFont = Object.prototype.hasOwnProperty.call(token, 'ff');
    const resolvedFill = hasFill ? resolveTokenColor(token.fb, SN) : null;

    return {
        hasFormat,
        // Token with format but no fill means explicit "no fill" (keep white sheet background).
        bgColor: hasFill ? resolvedFill : null,
        fgColor: hasFont ? resolveTokenColor(token.ff, SN) : undefined,
        isBold: token.b === 1 ? true : undefined,
        stripeSize: token.s || 1
    };
}

/**
 * Resolve a table style with workbook theme tokens (Excel-like).
 * Falls back to TABLE_STYLES when token/theme data is unavailable.
 */
export function resolveTableStyle(styleId, SN) {
    if (styleId === null) return null;

    const fallback = TABLE_STYLES[styleId] || TABLE_STYLES.TableStyleMedium2;
    const tokens = TABLE_STYLE_TOKENS[styleId];
    if (!tokens) return { ...fallback, _fromTokens: false };

    const whole = resolveElement(tokens, STYLE_ELEMENT.wholeTable, SN);
    const header = resolveElement(tokens, STYLE_ELEMENT.headerRow, SN);
    const totals = resolveElement(tokens, STYLE_ELEMENT.totalRow, SN);
    const first = resolveElement(tokens, STYLE_ELEMENT.firstColumn, SN);
    const last = resolveElement(tokens, STYLE_ELEMENT.lastColumn, SN);
    const rowStripe1 = resolveElement(tokens, STYLE_ELEMENT.rowStripe1, SN);
    const rowStripe2 = resolveElement(tokens, STYLE_ELEMENT.rowStripe2, SN);
    const colStripe1 = resolveElement(tokens, STYLE_ELEMENT.columnStripe1, SN);
    const colStripe2 = resolveElement(tokens, STYLE_ELEMENT.columnStripe2, SN);

    const rowStripe1Bg = pickDefined(rowStripe1.bgColor, whole.bgColor, fallback.rowStripe1);
    const rowStripe2Bg = pickDefined(rowStripe2.bgColor, whole.bgColor, fallback.rowStripe2);

    return {
        _fromTokens: true,
        name: fallback.name || styleId,
        headerBg: pickDefined(header.bgColor, whole.bgColor, fallback.headerBg),
        headerFg: pickDefined(header.fgColor, whole.fgColor, fallback.headerFg),
        headerBold: pickDefined(header.isBold, true),
        headerBorder: fallback.headerBorder,

        dataBg: pickDefined(whole.bgColor, fallback.dataBg, null),
        dataFg: pickDefined(whole.fgColor, fallback.dataFg, null),
        dataBold: pickDefined(fallback.dataBold, false),

        rowStripe1: rowStripe1Bg,
        rowStripe2: rowStripe2Bg,
        rowStripe1Size: rowStripe1.stripeSize || 1,
        rowStripe2Size: rowStripe2.stripeSize || rowStripe1.stripeSize || 1,

        columnStripe1: pickDefined(colStripe1.bgColor, fallback.columnStripe1, undefined),
        columnStripe2: pickDefined(colStripe2.bgColor, fallback.columnStripe2, undefined),
        columnStripe1Size: colStripe1.stripeSize || 1,
        columnStripe2Size: colStripe2.stripeSize || colStripe1.stripeSize || 1,

        firstColumnBg: first.hasFormat ? pickDefined(first.bgColor, whole.bgColor, null) : fallback.firstColumnBg,
        firstColumnFg: pickDefined(first.fgColor, whole.fgColor, fallback.firstColumnFg, fallback.dataFg, null),
        firstColumnBold: pickDefined(first.isBold, fallback.firstColumnBold, false),

        lastColumnBg: last.hasFormat ? pickDefined(last.bgColor, whole.bgColor, null) : fallback.lastColumnBg,
        lastColumnFg: pickDefined(last.fgColor, whole.fgColor, fallback.lastColumnFg, fallback.dataFg, null),
        lastColumnBold: pickDefined(last.isBold, fallback.lastColumnBold, false),

        totalsBg: pickDefined(totals.bgColor, whole.bgColor, fallback.totalsBg),
        totalsFg: pickDefined(totals.fgColor, whole.fgColor, fallback.totalsFg, fallback.dataFg, null),
        totalsBold: pickDefined(totals.isBold, true),

        border: fallback.border
    };
}

export const TABLE_STYLE_GROUPS = {
    light: Object.keys(TABLE_STYLES).filter(k => k.startsWith('TableStyleLight')),
    medium: Object.keys(TABLE_STYLES).filter(k => k.startsWith('TableStyleMedium')),
    dark: Object.keys(TABLE_STYLES).filter(k => k.startsWith('TableStyleDark'))
};

export function getStylePreviewColors(styleId, SN) {
    const style = resolveTableStyle(styleId, SN) || TABLE_STYLES[styleId];
    if (!style) return null;

    return {
        header: style.headerBg || '#FFFFFF',
        stripe1: style.rowStripe1 || '#FFFFFF',
        stripe2: style.rowStripe2 || '#F2F2F2',
        border: style.border || '#D9D9D9'
    };
}
