import enUS from './en-US.js';

function _isObject(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
}

function _clone(value) {
    if (Array.isArray(value)) return value.map(_clone);
    if (_isObject(value)) {
        const out = {};
        for (const key of Object.keys(value)) out[key] = _clone(value[key]);
        return out;
    }
    return value;
}

function _mergeDeep(target, source) {
    if (!_isObject(target) || !_isObject(source)) return target;
    for (const key of Object.keys(source)) {
        const sourceVal = source[key];
        const targetVal = target[key];
        if (_isObject(sourceVal) && _isObject(targetVal)) {
            _mergeDeep(targetVal, sourceVal);
            continue;
        }
        target[key] = sourceVal;
    }
    return target;
}

/** @type {Map<string, Object>} */
const _packs = new Map();
_packs.set('en-US', _clone(enUS));

/** @param {string} locale @param {Object} pack */
export function registerLocale(locale, pack) {
    if (!locale || !_isObject(pack)) return;
    const existing = _packs.get(locale);
    if (!existing) {
        _packs.set(locale, _clone(pack));
        return;
    }
    _mergeDeep(existing, _clone(pack));
}

/** @param {string} locale @returns {Object|undefined} */
export function getLocalePack(locale) {
    const pack = _packs.get(locale);
    return pack ? _clone(pack) : undefined;
}

/** @returns {Record<string, Object>} */
export function getAllLocalePacks() {
    const out = {};
    for (const [locale, pack] of _packs.entries()) out[locale] = _clone(pack);
    return out;
}
