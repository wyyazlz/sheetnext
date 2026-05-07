const DEFAULT_PRIMARY = '#2f6f4e';
const WHITE = { r: 255, g: 255, b: 255 };
const BLACK = { r: 0, g: 0, b: 0 };

const THEME_TOKENS = [
    'primary',
    'primaryRgb',
    'primaryHover',
    'primaryActive',
    'primaryLight',
    'primarySoft',
    'primarySoftHover',
    'primaryBorder',
    'primarySubtleBorder',
    'primaryRing',
    'primaryShadow',
    'primaryContrast'
];

function _kebabCase(value) {
    return String(value).replace(/[A-Z]/g, match => `-${match.toLowerCase()}`);
}

function _normalizeHex(value) {
    const text = String(value || '').trim();
    const match = text.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (!match) return null;
    const hex = match[1];
    if (hex.length === 3) {
        return `#${hex.split('').map(char => char + char).join('')}`.toLowerCase();
    }
    return `#${hex.toLowerCase()}`;
}

function _hexToRgb(value) {
    const hex = _normalizeHex(value);
    if (!hex) return null;
    const number = parseInt(hex.slice(1), 16);
    return {
        r: (number >> 16) & 255,
        g: (number >> 8) & 255,
        b: number & 255
    };
}

function _rgbToHex(rgb) {
    const toHex = (value) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0');
    return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`;
}

function _mix(base, target, weight) {
    return {
        r: base.r + (target.r - base.r) * weight,
        g: base.g + (target.g - base.g) * weight,
        b: base.b + (target.b - base.b) * weight
    };
}

function _rgba(rgb, alpha) {
    return `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${alpha})`;
}

function _coerceTheme(theme) {
    if (typeof theme === 'string') return { primary: theme };
    if (!theme || typeof theme !== 'object') return {};
    return {
        ...theme,
        ...(theme.colors && typeof theme.colors === 'object' ? theme.colors : {})
    };
}

export function createTheme(theme = {}) {
    const source = _coerceTheme(theme);
    const primary = _normalizeHex(source.primary) || DEFAULT_PRIMARY;
    const primaryRgb = _hexToRgb(primary) || _hexToRgb(DEFAULT_PRIMARY);
    const next = {
        primary,
        primaryRgb: `${primaryRgb.r}, ${primaryRgb.g}, ${primaryRgb.b}`,
        primaryHover: _rgbToHex(_mix(primaryRgb, BLACK, 0.16)),
        primaryActive: _rgbToHex(_mix(primaryRgb, BLACK, 0.28)),
        primaryLight: _rgbToHex(_mix(primaryRgb, WHITE, 0.34)),
        primarySoft: _rgbToHex(_mix(primaryRgb, WHITE, 0.92)),
        primarySoftHover: _rgbToHex(_mix(primaryRgb, WHITE, 0.84)),
        primaryBorder: _rgbToHex(_mix(primaryRgb, WHITE, 0.55)),
        primarySubtleBorder: _rgbToHex(_mix(primaryRgb, WHITE, 0.74)),
        primaryRing: _rgba(primaryRgb, 0.14),
        primaryShadow: _rgba(primaryRgb, 0.28),
        primaryContrast: '#ffffff'
    };

    THEME_TOKENS.forEach(token => {
        if (source[token]) next[token] = source[token];
    });

    return next;
}

export function applyTheme(root, theme = {}) {
    const colors = createTheme(theme);
    if (root?.style) {
        THEME_TOKENS.forEach(token => {
            root.style.setProperty(`--sn-${_kebabCase(token)}`, colors[token]);
        });
    }
    return colors;
}

export function getThemeColor(SN, token, fallback = '') {
    const root = SN?.containerDom || (typeof document !== 'undefined' ? document.documentElement : null);
    if (typeof window === 'undefined' || !root) return fallback;
    const value = window.getComputedStyle(root).getPropertyValue(`--sn-${_kebabCase(token)}`).trim();
    return value || fallback;
}
