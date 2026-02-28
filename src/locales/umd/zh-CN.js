import zhCN from '../zh-CN.js';

const root = typeof globalThis !== 'undefined' ? globalThis : (typeof window !== 'undefined' ? window : {});
const pendingKey = '__SHEETNEXT_PENDING_LOCALES__';

if (root.SheetNext && typeof root.SheetNext.registerLocale === 'function') {
    root.SheetNext.registerLocale('zh-CN', zhCN);
} else {
    if (!root[pendingKey]) root[pendingKey] = {};
    root[pendingKey]['zh-CN'] = zhCN;
}

export default zhCN;
