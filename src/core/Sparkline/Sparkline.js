import SparklineItem from './SparklineItem.js';

const DEFAULT_SPARKLINE_COLORS = {
    series: '#376092',
    negative: '#D00000',
    axis: '#000000',
    markers: '#D00000',
    first: '#D00000',
    last: '#D00000',
    high: '#D00000',
    low: '#D00000'
};

/**
 * 迷你图管理器
 * @title 📈 迷你图
 * 负责解析、存储和管理 Excel 迷你图数据
 * @class
 */
export default class Sparkline {
    /**
     * @param {Sheet} sheet - 所属工作表
     */
    constructor(sheet) {
        /**
         * 所属工作表
         * @type {Sheet}
         */
        this.sheet = sheet;
        /**
         * 按位置索引的迷你图 Map： "R:C" -> SparklineItem
         * @type {Map<string, SparklineItem>}
         */
        this.map = new Map();
        /**
         * 迷你图组列表
         * @type {Array<Object>}
         */
        this.groups = [];
    }

    // ==================== 增删改查 ====================

    /**
     * 添加迷你图
     * @param {Object} config - 配置对象
     * @param {string|Object} config.cellRef - 单元格引用 "A1" 或 {r, c}
     * @param {string} config.formula - 数据范围公式
     * @param {string} [config.type='line'] - 迷你图类型
     * @param {Object} [config.colors={}] - 颜色配置
     * @returns {SparklineItem|null}
     */
    add(config) {
        const SN = this.sheet.SN;
        // 解析 cellRef
        const cell = typeof config.cellRef === 'string'
            ? SN.Utils.cellStrToNum(config.cellRef)
            : config.cellRef;
        const r = cell.r;
        const c = cell.c;
        const { formula, type = 'line', colors = {} } = config;

        if (SN.Event.hasListeners('beforeSparklineAdd')) {
            const beforeEvent = SN.Event.emit('beforeSparklineAdd', { sheet: this.sheet, config });
            if (beforeEvent.canceled) return null;
        }

        // 创建或获取组
        let group = this.groups.find(g => g.type === type);
        if (!group) {
            group = this._createGroup(type, colors);
            this.groups.push(group);
        }

        const item = new SparklineItem({
            r,
            c,
            sqref: SN.Utils.cellNumToStr({ r, c }),
            formula,
            group
        }, this.sheet);

        this.map.set(item.key, item);

        if (SN.Event.hasListeners('afterSparklineAdd')) {
            SN.Event.emit('afterSparklineAdd', { sheet: this.sheet, config, sparkline: item });
        }
        return item;
    }

    /**
     * 删除迷你图
     * @param {string|Object} cellRef - 单元格引用 "A1" 或 {r, c}
     * @returns {boolean}
     */
    remove(cellRef) {
        const cell = typeof cellRef === 'string'
            ? this.sheet.SN.Utils.cellStrToNum(cellRef)
            : cellRef;
        const r = cell.r;
        const c = cell.c;
        const key = `${r}:${c}`;
        const item = this.map.get(key);
        if (!item) return false;

        const SN = this.sheet.SN;
        if (SN.Event.hasListeners('beforeSparklineRemove')) {
            const beforeEvent = SN.Event.emit('beforeSparklineRemove', {
                sheet: this.sheet,
                sparkline: item,
                r,
                c
            });
            if (beforeEvent.canceled) return false;
        }

        this.map.delete(key);

        if (SN.Event.hasListeners('afterSparklineRemove')) {
            SN.Event.emit('afterSparklineRemove', {
                sheet: this.sheet,
                sparkline: item,
                r,
                c
            });
        }
        return true;
    }

    /**
     * 根据单元格位置获取迷你图
     * @param {string|Object} cellRef - 单元格引用 "A1" 或 {r, c}
     * @returns {SparklineItem|null}
     */
    get(cellRef) {
        const cell = typeof cellRef === 'string'
            ? this.sheet.SN.Utils.cellStrToNum(cellRef)
            : cellRef;
        return this.map.get(`${cell.r}:${cell.c}`) ?? null;
    }

    /**
     * 获取所有迷你图实例
     * @returns {Array<SparklineItem>}
     */
    getAll() {
        return Array.from(this.map.values());
    }

    // ==================== 缓存管理 ====================

    /**
     * 清除指定迷你图的缓存
     * @param {string|Object|number} cellRefOrRow - 单元格引用 "A1"/{r,c} 或行索引
     * @param {number} [c] - 列索引（仅在第一个参数为行索引时使用）
     */
    clearCache(cellRefOrRow, c) {
        const cellRef = typeof cellRefOrRow === 'number'
            ? { r: cellRefOrRow, c }
            : cellRefOrRow;
        const item = this.get(cellRef);
        if (item) item._clearCache();
    }

    /**
     * 清除所有缓存
     */
    clearAllCache() {
        this.map.forEach(item => item._clearCache());
    }

    // ==================== 解析 ====================

    /**
     * 从 xmlObj 解析迷你图数据
     * @param {Object} xmlObj - Sheet 的 XML 对象
     */
    parse(xmlObj) {
        const extLst = xmlObj?.extLst;
        if (!extLst) return;

        const ext = Array.isArray(extLst.ext) ? extLst.ext : [extLst.ext];
        const sparklineExt = ext.find(e => e['_$uri'] === '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}');
        if (!sparklineExt) return;

        const sparklineGroups = sparklineExt['x14:sparklineGroups']?.['x14:sparklineGroup'];
        if (!sparklineGroups) return;

        const groups = Array.isArray(sparklineGroups) ? sparklineGroups : [sparklineGroups];

        groups.forEach((groupData, groupIndex) => {
            const group = this._parseGroup(groupData, groupIndex);
            this.groups.push(group);

            const sparklines = groupData['x14:sparklines']?.['x14:sparkline'];
            if (!sparklines) return;

            const instances = Array.isArray(sparklines) ? sparklines : [sparklines];
            instances.forEach(sparklineData => {
                const item = this._parseInstance(sparklineData, group);
                this.map.set(item.key, item);
            });
        });
    }

    /**
     * @private
     */
    _parseGroup(groupData, index) {
        const colors = { ...DEFAULT_SPARKLINE_COLORS };
        const colorMap = {
            series: 'x14:colorSeries',
            negative: 'x14:colorNegative',
            axis: 'x14:colorAxis',
            markers: 'x14:colorMarkers',
            first: 'x14:colorFirst',
            last: 'x14:colorLast',
            high: 'x14:colorHigh',
            low: 'x14:colorLow'
        };
        Object.entries(colorMap).forEach(([key, xmlKey]) => {
            const parsed = this._parseColor(groupData[xmlKey]);
            if (parsed) colors[key] = parsed;
        });

        return {
            id: index,
            type: groupData['_$type'] || 'line',
            lineWeight: groupData['_$lineWeight'] ? parseFloat(groupData['_$lineWeight']) : 0.75,
            manualMax: groupData['_$manualMax'] ? parseFloat(groupData['_$manualMax']) : null,
            manualMin: groupData['_$manualMin'] ? parseFloat(groupData['_$manualMin']) : null,
            dateAxis: groupData['_$dateAxis'] === '1',
            markers: groupData['_$markers'] === '1',
            high: groupData['_$high'] === '1',
            low: groupData['_$low'] === '1',
            first: groupData['_$first'] === '1',
            last: groupData['_$last'] === '1',
            negative: groupData['_$negative'] === '1',
            displayEmptyCellsAs: groupData['_$displayEmptyCellsAs'] || 'gap',
            colors
        };
    }

    /**
     * @private
     */
    _parseInstance(sparklineData, group) {
        const sqref = sparklineData['xm:sqref'];
        const formula = sparklineData['xm:f'];
        const pos = this.sheet.SN.Utils.cellStrToNum(sqref);

        return new SparklineItem({
            r: pos.r,
            c: pos.c,
            sqref,
            formula,
            group
        }, this.sheet);
    }

    /**
     * @private
     */
    _parseColor(colorObj) {
        if (!colorObj) return null;
        const rgb = colorObj['_$rgb'];
        if (!rgb) return null;
        return '#' + rgb.slice(2);
    }

    /**
     * @private
     */
    _createGroup(type, colors = {}) {
        return {
            id: this.groups.length,
            type,
            displayEmptyCellsAs: 'gap',
            colors: {
                series: colors.series || DEFAULT_SPARKLINE_COLORS.series,
                negative: colors.negative || DEFAULT_SPARKLINE_COLORS.negative,
                axis: colors.axis || DEFAULT_SPARKLINE_COLORS.axis,
                markers: colors.markers || DEFAULT_SPARKLINE_COLORS.markers,
                first: colors.first || DEFAULT_SPARKLINE_COLORS.first,
                last: colors.last || DEFAULT_SPARKLINE_COLORS.last,
                high: colors.high || DEFAULT_SPARKLINE_COLORS.high,
                low: colors.low || DEFAULT_SPARKLINE_COLORS.low
            },
            lineWeight: 0.75,
            manualMax: null,
            manualMin: null,
            dateAxis: false,
            markers: false,
            high: false,
            low: false,
            first: false,
            last: false,
            negative: false
        };
    }

    // ==================== 导出 ====================

    /**
     * 导出为 XML 对象结构
     * @returns {Object|null}
     */
    toXmlObj() {
        if (this.groups.length === 0) return null;

        const sparklineGroups = this.groups.map(group => {
            const instances = this.getAll().filter(item => item.group.id === group.id);
            if (instances.length === 0) return null;

            const groupObj = {
                'x14:colorSeries': { '_$rgb': 'FF' + group.colors.series.slice(1) },
                'x14:colorNegative': { '_$rgb': 'FF' + group.colors.negative.slice(1) },
                'x14:colorAxis': { '_$rgb': 'FF' + group.colors.axis.slice(1) },
                'x14:colorMarkers': { '_$rgb': 'FF' + group.colors.markers.slice(1) },
                'x14:colorFirst': { '_$rgb': 'FF' + group.colors.first.slice(1) },
                'x14:colorLast': { '_$rgb': 'FF' + group.colors.last.slice(1) },
                'x14:colorHigh': { '_$rgb': 'FF' + group.colors.high.slice(1) },
                'x14:colorLow': { '_$rgb': 'FF' + group.colors.low.slice(1) },
                'x14:sparklines': {
                    'x14:sparkline': instances.map(item => item._toXmlObject())
                },
                '_$displayEmptyCellsAs': group.displayEmptyCellsAs
            };

            if (group.type !== 'line') groupObj['_$type'] = group.type;
            if (group.lineWeight !== 0.75) groupObj['_$lineWeight'] = group.lineWeight.toString();
            if (group.manualMax !== null) groupObj['_$manualMax'] = group.manualMax.toString();
            if (group.manualMin !== null) groupObj['_$manualMin'] = group.manualMin.toString();
            if (group.dateAxis) groupObj['_$dateAxis'] = '1';
            if (group.markers) groupObj['_$markers'] = '1';
            if (group.high) groupObj['_$high'] = '1';
            if (group.low) groupObj['_$low'] = '1';
            if (group.first) groupObj['_$first'] = '1';
            if (group.last) groupObj['_$last'] = '1';
            if (group.negative) groupObj['_$negative'] = '1';

            return groupObj;
        }).filter(Boolean);

        if (sparklineGroups.length === 0) return null;

        return {
            'ext': [{
                'x14:sparklineGroups': {
                    'x14:sparklineGroup': sparklineGroups,
                    '_$xmlns:xm': 'http://schemas.microsoft.com/office/excel/2006/main'
                },
                '_$uri': '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}',
                '_$xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
            }]
        };
    }
}
