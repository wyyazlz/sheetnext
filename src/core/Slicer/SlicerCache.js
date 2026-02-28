import { buildSlicerKey, formatSlicerValue, isBlankValue } from './helpers.js';

export default class SlicerCache {
    constructor(SN, options = {}) {
        this.cacheId = options.cacheId ?? this._generateCacheId(SN);
        this.sourceType = options.sourceType ?? 'table';
        this.sourceSheet = options.sourceSheet ?? null;
        this.sourceName = options.sourceName ?? null;
        this.sourceId = options.sourceId ?? null;
        this.fieldIndex = Number.isFinite(options.fieldIndex) ? options.fieldIndex : 0;
        this.fieldName = options.fieldName ?? '';
        this.items = [];
        this._SN = SN;
        this._itemMap = new Map();
        this._version = 0;
        this._dirty = true;
    }

    _generateCacheId(SN = this._SN) {
        const existing = SN?._slicerCaches ? Array.from(SN._slicerCaches.keys()) : [];
        let maxId = 0;
        existing.forEach(id => {
            if (id > maxId) maxId = id;
        });
        return maxId + 1;
    }

    get version() {
        return this._version;
    }

    markDirty() {
        this._dirty = true;
    }

    getItemMap() {
        return this._itemMap;
    }

    getItems({ force = false } = {}) {
        if (this._dirty || force) {
            this.refresh();
        }
        return this.items;
    }

    getSourceSheet() {
        if (!this.sourceSheet) return null;
        return this._SN.getSheet(this.sourceSheet);
    }

    getSourceTable() {
        const sheet = this.getSourceSheet();
        if (!sheet?.Table) return null;
        if (this.sourceId) {
            const byId = sheet.Table.get(this.sourceId);
            if (byId) return byId;
        }
        if (this.sourceName) {
            return sheet.Table.getByName(this.sourceName);
        }
        return null;
    }

    getSourcePivotTable() {
        const sheet = this.getSourceSheet();
        if (!sheet?.PivotTable || !this.sourceName) return null;
        return sheet.PivotTable.get(this.sourceName) || null;
    }

    refresh() {
        if (this.sourceType === 'pivot') {
            this._refreshFromPivot();
        } else {
            this._refreshFromTable();
        }
        this._dirty = false;
        this._version += 1;
    }

    _refreshFromTable() {
        const table = this.getSourceTable();
        if (!table?.range) {
            this.items = [];
            this._itemMap = new Map();
            return;
        }

        const sheet = table.sheet;
        const range = table.range;
        const colIndex = range.s.c + this.fieldIndex;
        const startRow = range.s.r + (table.showHeaderRow ? 1 : 0);
        const endRow = table.showTotalsRow ? range.e.r - 1 : range.e.r;

        const items = [];
        const itemMap = new Map();

        for (let r = startRow; r <= endRow; r++) {
            const cell = sheet.getCell(r, colIndex);
            const value = cell.calcVal ?? cell.editVal;
            const key = buildSlicerKey(value);
            if (itemMap.has(key)) {
                itemMap.get(key).count += 1;
                continue;
            }
            const item = {
                key,
                value: isBlankValue(value) ? null : value,
                text: formatSlicerValue(value),
                count: 1
            };
            itemMap.set(key, item);
            items.push(item);
        }

        this.items = items;
        this._itemMap = itemMap;
    }

    _refreshFromPivot() {
        const pivotTable = this.getSourcePivotTable();
        const cache = pivotTable?.cache;
        const cacheField = cache?.fields?.[this.fieldIndex];
        const sharedItems = cacheField?.sharedItems?.values ?? [];

        const items = [];
        const itemMap = new Map();

        sharedItems.forEach((sharedItem, sharedIndex) => {
            const value = sharedItem?.value ?? null;
            const key = buildSlicerKey(value);
            if (itemMap.has(key)) return;
            const text = pivotTable?._getItemDisplayValue?.(this.fieldIndex, sharedIndex) ?? formatSlicerValue(value);
            const item = {
                key,
                value,
                text,
                sharedIndex,
                count: 0
            };
            itemMap.set(key, item);
            items.push(item);
        });

        this.items = items;
        this._itemMap = itemMap;
    }
}
