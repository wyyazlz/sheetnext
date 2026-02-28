/**
 * 公式目录模块
 * 提供公式分类、搜索、快捷访问等功能
 */
import { FORMULA_DATA } from '../../assets/formulaData.js';

// 分类定义
const CATEGORIES = [
    { key: 'financial', labelKey: 'action.formula.content.categories.financial' },
    { key: 'logical', labelKey: 'action.formula.content.categories.logical' },
    { key: 'text', labelKey: 'action.formula.content.categories.text' },
    { key: 'datetime', labelKey: 'action.formula.content.categories.datetime' },
    { key: 'lookup', labelKey: 'action.formula.content.categories.lookup' },
    { key: 'math', labelKey: 'action.formula.content.categories.math' },
    { key: 'statistical', labelKey: 'action.formula.content.categories.statistical' },
    { key: 'engineering', labelKey: 'action.formula.content.categories.engineering' },
    { key: 'information', labelKey: 'action.formula.content.categories.information' },
    { key: 'compatibility', labelKey: 'action.formula.content.categories.compatibility' },
    { key: 'database', labelKey: 'action.formula.content.categories.database' },
    { key: 'web', labelKey: 'action.formula.content.categories.web' },
    { key: 'cube', labelKey: 'action.formula.content.categories.cube' },
    { key: 'addins', labelKey: 'action.formula.content.categories.addins' }
];

const DIALOG_CATEGORIES = [
    { key: 'all', labelKey: 'action.formula.content.categories.all' },
    { key: 'recent', labelKey: 'action.formula.content.categories.recent' },
    ...CATEGORIES
];

// 常用公式快捷方式
const CATEGORY_SHORTCUTS = {
    common: ['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'IF', 'AND', 'OR', 'VLOOKUP', 'XLOOKUP', 'TEXT', 'DATE', 'TODAY'],
    financial: ['PMT', 'PV', 'FV', 'NPV', 'IRR', 'RATE', 'NPER', 'IPMT', 'PPMT', 'XIRR', 'XNPV', 'SLN'],
    logical: ['IF', 'IFS', 'AND', 'OR', 'NOT', 'XOR', 'SWITCH', 'IFERROR', 'IFNA', 'LET', 'LAMBDA'],
    text: ['CONCAT', 'TEXT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'FIND', 'SEARCH', 'REPLACE', 'SUBSTITUTE', 'TRIM', 'UPPER', 'LOWER'],
    datetime: ['TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'EDATE', 'EOMONTH', 'WORKDAY', 'NETWORKDAYS', 'DATEDIF'],
    lookup: ['XLOOKUP', 'VLOOKUP', 'HLOOKUP', 'INDEX', 'MATCH', 'XMATCH', 'OFFSET', 'INDIRECT', 'CHOOSE', 'LOOKUP', 'FILTER', 'SORT', 'UNIQUE'],
    math: ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'ABS', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'MOD', 'POWER', 'SQRT', 'RAND', 'RANDBETWEEN']
};

const RECENT_KEY = 'sn-formula-recent';
const RECENT_LIMIT = 20;

// 构建索引
const CATEGORY_LABEL_KEY_MAP = new Map(CATEGORIES.map(c => [c.key, c.labelKey]));
const FORMULA_MAP = new Map();
const FORMULAS_BY_CATEGORY = new Map();

CATEGORIES.forEach(c => FORMULAS_BY_CATEGORY.set(c.key, []));

FORMULA_DATA.forEach(item => {
    const entry = {
        name: item.name,
        categoryKey: item.category,
        categoryLabelKey: CATEGORY_LABEL_KEY_MAP.get(item.category) || '',
        desc: item.desc,
        version: item.version || ''
    };
    FORMULA_MAP.set(item.name, entry);
    const list = FORMULAS_BY_CATEGORY.get(item.category);
    if (list) list.push(entry);
});

const FORMULA_LIST = [...FORMULA_MAP.values()];

// 最近使用
function readRecentNames() {
    if (typeof localStorage === 'undefined') return [];
    try {
        const raw = localStorage.getItem(RECENT_KEY);
        const arr = JSON.parse(raw || '[]');
        return Array.isArray(arr) ? arr : [];
    } catch {
        return [];
    }
}

function writeRecentNames(names) {
    if (typeof localStorage === 'undefined') return;
    localStorage.setItem(RECENT_KEY, JSON.stringify(names.slice(0, RECENT_LIMIT)));
}

function getRecentFormulas() {
    return readRecentNames().map(name => FORMULA_MAP.get(name)).filter(Boolean);
}

export const FormulaCatalog = {
    getDialogCategories() {
        return DIALOG_CATEGORIES;
    },

    getFormula(name) {
        if (!name) return null;
        return FORMULA_MAP.get(name.toUpperCase()) || null;
    },

    getFormulasByCategory(categoryKey) {
        if (categoryKey === 'recent') return getRecentFormulas();
        if (categoryKey === 'all') return FORMULA_LIST;
        return FORMULAS_BY_CATEGORY.get(categoryKey) || [];
    },

    search(keyword, categoryKey = 'all') {
        const list = this.getFormulasByCategory(categoryKey);
        const term = (keyword || '').trim().toLowerCase();
        if (!term) return list;
        return list.filter(item =>
            item.name.toLowerCase().includes(term) ||
            (item.desc || '').toLowerCase().includes(term)
        );
    },

    addRecent(name) {
        const finalName = name.toUpperCase();
        if (!FORMULA_MAP.has(finalName)) return;
        const names = readRecentNames().filter(item => item !== finalName);
        names.unshift(finalName);
        writeRecentNames(names);
    },

    getShortcutFormulas(categoryKey, limit = 12) {
        const shortcuts = CATEGORY_SHORTCUTS[categoryKey] || [];
        const picked = [];
        shortcuts.forEach(name => {
            const formula = FORMULA_MAP.get(name);
            if (formula && !picked.includes(formula)) picked.push(formula);
        });
        if (picked.length < limit) {
            const list = categoryKey === 'common' ? FORMULA_LIST : this.getFormulasByCategory(categoryKey);
            list.forEach(formula => {
                if (picked.length >= limit) return;
                if (!picked.includes(formula)) picked.push(formula);
            });
        }
        return picked.slice(0, limit);
    }
};

export default FormulaCatalog;
