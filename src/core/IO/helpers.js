/**
 * IO 辅助函数和工具类
 */

// ==================== IndexTable 类 - 用于去重和快速查找 ====================

function _stableSerialize(value) {
    if (value === null) return 'null';
    if (value === undefined) return 'undefined';

    if (value instanceof Date) {
        return `date:${value.toISOString()}`;
    }

    const valueType = typeof value;
    if (valueType === 'number') {
        if (Number.isNaN(value)) return 'number:NaN';
        if (!Number.isFinite(value)) return `number:${value > 0 ? 'Infinity' : '-Infinity'}`;
        return `number:${value}`;
    }
    if (valueType === 'string') return `string:${JSON.stringify(value)}`;
    if (valueType === 'boolean') return `boolean:${value}`;
    if (valueType !== 'object') return `${valueType}:${String(value)}`;

    if (Array.isArray(value)) {
        return `array:[${value.map(item => _stableSerialize(item)).join(',')}]`;
    }

    const keys = Object.keys(value).sort();
    const body = keys.map(key => `${JSON.stringify(key)}:${_stableSerialize(value[key])}`).join(',');
    return `object:{${body}}`;
}

export class IndexTable {
    constructor(hashFn) {
        this.items = [];
        this.map = new Map();
        this.hashFn = hashFn || this._defaultHash;
    }

    add(item) {
        if (item === null || item === undefined) return null;
        const key = this.hashFn(item);
        if (this.map.has(key)) return this.map.get(key);
        const index = this.items.length;
        this.items.push(item);
        this.map.set(key, index);
        return index;
    }

    _defaultHash(obj) {
        return _stableSerialize(obj);
    }

    getAll() {
        return this.items;
    }
}

// ==================== 辅助函数 ====================

// 判断两个单元格数据是否相等（用于行程编码）
export function isCellDataEqual(cell1, cell2) {
    const keys = ['v', 'vt', 'f', 'cv', 'cvt', 's', 'fo', 'fi', 'b', 'a', 'fmt'];
    for (const key of keys) {
        if (cell1[key] !== cell2[key]) return false;
    }
    if (cell1.link || cell2.link) {
        if (!cell1.link || !cell2.link) return false;
        if (cell1.link.t !== cell2.link.t || cell1.link.l !== cell2.link.l) return false;
    }
    if (_stableSerialize(cell1.dv) !== _stableSerialize(cell2.dv)) return false;
    if (_stableSerialize(cell1.p) !== _stableSerialize(cell2.p)) return false;
    if (_stableSerialize(cell1.rt) !== _stableSerialize(cell2.rt)) return false;
    return true;
}

// 判断字符串是否应该转为数字
export function shouldConvertToNumber(str) {
    if (str.trim() === '') return false;
    if (/^0\d+/.test(str)) return false;
    const num = Number(str);
    return !isNaN(num) && isFinite(num);
}
