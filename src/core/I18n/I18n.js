/** @typedef {Record<string, any>} LocalePack */

function _flattenPack(pack, prefix = '', out = {}) {
    if (!pack || typeof pack !== 'object' || Array.isArray(pack)) return out;
    for (const key of Object.keys(pack)) {
        const value = pack[key];
        const nextKey = prefix ? `${prefix}.${key}` : key;
        if (value && typeof value === 'object' && !Array.isArray(value)) {
            _flattenPack(value, nextKey, out);
            continue;
        }
        if (typeof value === 'string') out[nextKey] = value;
    }
    return out;
}

/** Lightweight i18n service with flat key lookup. */
export default class I18n {
    /** @param {string} [locale='en-US'] @param {string} [fallbackLocale='en-US'] */
    constructor(locale = 'en-US', fallbackLocale = 'en-US') {
        /** @type {string} */
        this.locale = locale || 'en-US';
        /** @type {string} */
        this.fallbackLocale = fallbackLocale || 'en-US';
        /** @type {Map<string, Record<string, string>>} */
        this._packs = new Map();
    }

    /** @param {string} locale @param {LocalePack} pack */
    register(locale, pack) {
        if (!locale) return;
        const current = this._packs.get(locale) || {};
        const flatPack = _flattenPack(pack);
        this._packs.set(locale, { ...current, ...flatPack });
    }

    /** @param {string} locale */
    setLocale(locale) {
        if (!locale || typeof locale !== 'string') return;
        this.locale = locale;
    }

    /** @param {string} key @param {Record<string, any>} [params={}] @returns {string} */
    t(key, params = {}) {
        if (!key || typeof key !== 'string') return '';
        const localePack = this._packs.get(this.locale) || {};
        const fallbackPack = this._packs.get(this.fallbackLocale) || {};
        const template = localePack[key] ?? fallbackPack[key] ?? key;
        return String(template).replace(/\{(\w+)\}/g, (_, paramKey) => {
            if (Object.prototype.hasOwnProperty.call(params, paramKey)) return String(params[paramKey]);
            return `{${paramKey}}`;
        });
    }
}
