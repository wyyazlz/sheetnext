import SlicerItem from './SlicerItem.js';
import SlicerCache from './SlicerCache.js';

const DEFAULT_SLICER_COL_WIDTH = 2.6;
const DEFAULT_SLICER_ROW_HEIGHT = 11;

/**
 * 切片器管理器
 * 负责切片器的增删改查
 * @class
 */
export default class Slicer {
    /** @param {Sheet} sheet */
    constructor(sheet) {
        /** @type {Sheet} */
        this.sheet = sheet;
        /** @type {Map<string, SlicerItem>} */
        this.map = new Map();
        this._list = [];
    }

    /** @type {number} */
    get size() {
        return this.map.size;
    }

    /**
     * 添加切片器
     * @param {Object} config - 配置对象
     * @returns {Slicer}
     */
    add(config) {
        const SN = this.sheet.SN;

        if (SN.Event.hasListeners('beforeSlicerAdd')) {
            const beforeEvent = SN.Event.emit('beforeSlicerAdd', { sheet: this.sheet, config });
            if (beforeEvent.canceled) return null;
        }

        // 获取或创建缓存
        const cache = this._getOrCreateCache(config);
        cache.markDirty();

        // 生成 ID
        const id = this._buildId();

        // 计算默认尺寸和位置
        const { width, height } = this._getDefaultSize();
        const startCell = config.startCell ?? this.sheet.activeCell;

        // 创建 Drawing
        const drawing = this.sheet.Drawing._add({
            type: 'slicer',
            startCell,
            offsetX: config.offsetX || 0,
            offsetY: config.offsetY || 0,
            width: config.width || width,
            height: config.height || height,
            // Excel slicer default: move with cells but keep size.
            anchorType: config.anchorType || 'oneCell'
        });

        // 创建 SlicerItem 实例
        const slicer = new SlicerItem(this.sheet, {
            id,
            name: config.name || id,
            caption: config.caption || config.fieldName,
            cacheId: cache.cacheId,
            sourceType: config.sourceType || 'table',
            sourceSheet: config.sourceSheet || this.sheet.name,
            sourceName: config.sourceName,
            sourceId: config.sourceId,
            fieldIndex: config.fieldIndex,
            fieldName: config.fieldName,
            columns: config.columns,
            buttonHeight: config.buttonHeight,
            showHeader: config.showHeader,
            multiSelect: config.multiSelect,
            styleName: config.styleName,
            drawing
        });

        this.map.set(slicer.id, slicer);
        this._list.push(slicer);

        if (SN.Event.hasListeners('afterSlicerAdd')) {
            SN.Event.emit('afterSlicerAdd', { sheet: this.sheet, config, slicer });
        }

        return slicer;
    }

    /**
     * 内部添加（解析时使用，不触发事件）
     * @param {Slicer} slicer
     */
    _addParsed(slicer) {
        this.map.set(slicer.id, slicer);
        this._list.push(slicer);
    }

    /**
     * 删除切片器
     * @param {string} id
     */
    remove(id) {
        const slicer = this.map.get(id);
        if (!slicer) return;

        const SN = this.sheet.SN;
        if (SN.Event.hasListeners('beforeSlicerRemove')) {
            const beforeEvent = SN.Event.emit('beforeSlicerRemove', { sheet: this.sheet, slicer, id });
            if (beforeEvent.canceled) return;
        }

        // 删除关联的 Drawing
        if (slicer.Drawing) {
            this.sheet.Drawing.remove(slicer.Drawing.id);
        }

        // 清理 DOM
        if (slicer._dom) {
            slicer._dom.remove();
            slicer._dom = null;
        }

        this.map.delete(id);
        const idx = this._list.indexOf(slicer);
        if (idx > -1) this._list.splice(idx, 1);

        if (SN.Event.hasListeners('afterSlicerRemove')) {
            SN.Event.emit('afterSlicerRemove', { sheet: this.sheet, slicer, id });
        }
    }

    /**
     * 获取切片器
     * @param {string} id
     * @returns {Slicer|null}
     */
    get(id) {
        return this.map.get(id) || null;
    }

    /**
     * 获取所有切片器
     * @returns {Array<Slicer>}
     */
    getAll() {
        return this._list;
    }

    /**
     * 遍历所有切片器
     * @param {Function} callback
     */
    forEach(callback) {
        this._list.forEach(callback);
    }

    _buildId() {
        const SN = this.sheet.SN;
        if (!SN._slicerIdCounter) SN._slicerIdCounter = 1;
        return `Slicer_${SN._slicerIdCounter++}`;
    }

    _getDefaultSize() {
        const colWidth = this.sheet.defaultColWidth || 70;
        const rowHeight = this.sheet.defaultRowHeight || 20;
        return {
            width: Math.round(colWidth * DEFAULT_SLICER_COL_WIDTH),
            height: Math.round(rowHeight * DEFAULT_SLICER_ROW_HEIGHT)
        };
    }

    _getOrCreateCache(config) {
        const SN = this.sheet.SN;
        if (!SN._slicerCaches) {
            SN._slicerCaches = new Map();
        }

        const match = Array.from(SN._slicerCaches.values()).find(cache =>
            cache.sourceType === config.sourceType &&
            cache.sourceSheet === (config.sourceSheet || this.sheet.name) &&
            cache.sourceName === config.sourceName &&
            cache.fieldIndex === config.fieldIndex
        );

        if (match) return match;

        const cache = new SlicerCache(SN, {
            sourceType: config.sourceType || 'table',
            sourceSheet: config.sourceSheet || this.sheet.name,
            sourceName: config.sourceName,
            sourceId: config.sourceId,
            fieldIndex: config.fieldIndex,
            fieldName: config.fieldName
        });
        SN._slicerCaches.set(cache.cacheId, cache);
        return cache;
    }
}
