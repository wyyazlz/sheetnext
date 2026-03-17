// ==================== 颜色处理函数 ====================

/**
 * Get matching colours from xmlObj color
 * @param {Object} obj - Colour Object
 * @param {Object} SN - Global SN Object
 * @returns {string} - Hexadecimal colour value
 */
export function getColor(obj, SN) {
    if (typeof obj != 'object' || !obj) return undefined;
    if (obj['_$rgb']) return '#' + obj['_$rgb'].substring(2);
    if (obj['_$lastClr']) return '#' + obj['_$lastClr'];
    if (obj['_$indexed']) return getColorByIndex(Number(obj['_$indexed']));
    if (obj['_$theme']) return getThemeColor(parseFloat(obj['_$theme']), parseFloat(obj['_$tint'] ?? 0), SN);
    if (obj['_$val']) {
        if (/^[0-9A-Fa-f]{6}$/.test(obj['_$val'])) {
            return '#' + obj['_$val'].toUpperCase();
        }
        const schemeMap = {
            'lt1': 0,
            'dk1': 1,
            'lt2': 2,
            'dk2': 3,
            'bg1': 0,
            'tx1': 1,
            'bg2': 2,
            'tx2': 3,
            'accent1': 4,
            'accent2': 5,
            'accent3': 6,
            'accent4': 7,
            'accent5': 8,
            'accent6': 9,
            'hlink': 10,
            'folHlink': 11
        };
        const themeIndex = schemeMap[obj['_$val']];
        if (themeIndex !== undefined && SN) {
            let tint = 0;
            if (obj['a:tint']?._$val) tint = parseFloat(obj['a:tint']._$val) / 100000;
            else if (obj['a:shade']?._$val) tint = -parseFloat(obj['a:shade']._$val) / 100000;
            const baseColor = getThemeColor(themeIndex, tint, SN);
            const lumMod = parseFloat(obj['a:lumMod']?._$val);
            const lumOff = parseFloat(obj['a:lumOff']?._$val);
            if (Number.isFinite(lumMod) || Number.isFinite(lumOff)) {
                return applyLumToThemeColor(baseColor, lumMod, lumOff);
            }
            return baseColor;
        }
        return '#000000';
    }
    return '#000000';
}

/**
 * Get the corresponding colour values from the index
 * @param {number} index - Colour Index
 * @returns {string} - Hexadecimal colour value
 */
export function getColorByIndex(index) {
    const indexedColors = [
        '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
        '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
        '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
        '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
        '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
        '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
        '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
        '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'
    ];
    return indexedColors[index] ?? '#000000';
}

/**
 * Retrieving theme colours by theme index and color
 * @param {number} themeIndex - Theme Index
 * @param {number} tint - Hue
 * @param {Object} SN - Global SN Object
 * @returns {string} - Hexadecimal colour value
 */
export function getThemeColor(themeIndex, tint, SN) {
    if (!SN?.Xml) return '#000000';
    const themeobj = SN.Xml.clrScheme ?? SN.Xml.clrSchemeTp;
    const themeKeys = ["a:lt1", "a:dk1", "a:lt2", "a:dk2", "a:accent1", "a:accent2", "a:accent3", "a:accent4", "a:accent5", "a:accent6", "a:hlink", "a:folHlink"];
    const themeKey = themeKeys[themeIndex];
    const colorObj = themeobj[themeKey];
    if (!colorObj) throw new Error("Invalid theme index");
    let color;
    if (colorObj["a:srgbClr"]) color = colorObj["a:srgbClr"]["_$val"];
    else if (colorObj["a:sysClr"]) color = colorObj["a:sysClr"]["_$lastClr"];
    if (!color) throw new Error("Invalid color structure");
    return '#' + applyTintToColor(color, tint);
}

/**
 * Apply color changes to colours
 * @param {string} color - RGB colour string (excluding#)
 * @param {number} tint - Hue (-1 to 1)
 * @returns {string} - Processed RGB colour string
 */
export function applyTintToColor(color, tint) {
    const RGB_MAX = 255;
    const HLS_MAX = 240;
    const clamp = (value, min, max) => Math.min(max, Math.max(min, value));

    const baseR = parseInt(color.slice(0, 2), 16) / RGB_MAX;
    const baseG = parseInt(color.slice(2, 4), 16) / RGB_MAX;
    const baseB = parseInt(color.slice(4, 6), 16) / RGB_MAX;

    const max = Math.max(baseR, baseG, baseB);
    const min = Math.min(baseR, baseG, baseB);
    const delta = max - min;

    let hue = 0;
    let sat = 0;
    let lum = (max + min) / 2;

    if (delta > 0) {
        sat = lum > 0.5
            ? delta / (2 - max - min)
            : delta / (max + min);

        if (max === baseR) {
            hue = (baseG - baseB) / delta + (baseG < baseB ? 6 : 0);
        } else if (max === baseG) {
            hue = (baseB - baseR) / delta + 2;
        } else {
            hue = (baseR - baseG) / delta + 4;
        }
        hue /= 6;
    }

    let msLum = lum * HLS_MAX;
    if (tint < 0) {
        msLum = msLum * (1 + tint);
    } else {
        msLum = msLum * (1 - tint) + (HLS_MAX - HLS_MAX * (1 - tint));
    }
    lum = clamp(msLum / HLS_MAX, 0, 1);

    const hueToRgb = (p, q, t) => {
        let wrapped = t;
        if (wrapped < 0) wrapped += 1;
        if (wrapped > 1) wrapped -= 1;
        if (wrapped < 1 / 6) return p + (q - p) * 6 * wrapped;
        if (wrapped < 1 / 2) return q;
        if (wrapped < 2 / 3) return p + (q - p) * (2 / 3 - wrapped) * 6;
        return p;
    };

    let outR;
    let outG;
    let outB;
    if (sat === 0) {
        outR = lum;
        outG = lum;
        outB = lum;
    } else {
        const q = lum < 0.5
            ? lum * (1 + sat)
            : lum + sat - lum * sat;
        const p = 2 * lum - q;
        outR = hueToRgb(p, q, hue + 1 / 3);
        outG = hueToRgb(p, q, hue);
        outB = hueToRgb(p, q, hue - 1 / 3);
    }

    const r = clamp(Math.round(outR * RGB_MAX), 0, RGB_MAX);
    const g = clamp(Math.round(outG * RGB_MAX), 0, RGB_MAX);
    const b = clamp(Math.round(outB * RGB_MAX), 0, RGB_MAX);

    return ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

function applyLumToThemeColor(color, lumMod, lumOff) {
    const normalized = String(color || '').replace(/^#/, '');
    if (!/^[0-9A-Fa-f]{6}$/.test(normalized)) return color;

    const RGB_MAX = 255;
    const HLS_MAX = 240;
    const clamp = (value, min, max) => Math.min(max, Math.max(min, value));

    const baseR = parseInt(normalized.slice(0, 2), 16) / RGB_MAX;
    const baseG = parseInt(normalized.slice(2, 4), 16) / RGB_MAX;
    const baseB = parseInt(normalized.slice(4, 6), 16) / RGB_MAX;

    const max = Math.max(baseR, baseG, baseB);
    const min = Math.min(baseR, baseG, baseB);
    const delta = max - min;

    let hue = 0;
    let sat = 0;
    let lum = (max + min) / 2;

    if (delta > 0) {
        sat = lum > 0.5 ? delta / (2 - max - min) : delta / (max + min);
        if (max === baseR) hue = (baseG - baseB) / delta + (baseG < baseB ? 6 : 0);
        else if (max === baseG) hue = (baseB - baseR) / delta + 2;
        else hue = (baseR - baseG) / delta + 4;
        hue /= 6;
    }

    let msLum = lum * HLS_MAX;
    const mod = Number.isFinite(lumMod) ? lumMod / 100000 : 1;
    const off = Number.isFinite(lumOff) ? lumOff / 100000 : 0;
    msLum = clamp(msLum * mod + HLS_MAX * off, 0, HLS_MAX);
    lum = msLum / HLS_MAX;

    const hueToRgb = (p, q, t) => {
        let wrapped = t;
        if (wrapped < 0) wrapped += 1;
        if (wrapped > 1) wrapped -= 1;
        if (wrapped < 1 / 6) return p + (q - p) * 6 * wrapped;
        if (wrapped < 1 / 2) return q;
        if (wrapped < 2 / 3) return p + (q - p) * (2 / 3 - wrapped) * 6;
        return p;
    };

    let outR;
    let outG;
    let outB;
    if (sat === 0) {
        outR = lum;
        outG = lum;
        outB = lum;
    } else {
        const q = lum < 0.5 ? lum * (1 + sat) : lum + sat - lum * sat;
        const p = 2 * lum - q;
        outR = hueToRgb(p, q, hue + 1 / 3);
        outG = hueToRgb(p, q, hue);
        outB = hueToRgb(p, q, hue - 1 / 3);
    }

    const r = clamp(Math.round(outR * RGB_MAX), 0, RGB_MAX);
    const g = clamp(Math.round(outG * RGB_MAX), 0, RGB_MAX);
    const b = clamp(Math.round(outB * RGB_MAX), 0, RGB_MAX);
    return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}
