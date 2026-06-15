import {
    CHART_CATEGORIES,
    getChartDisplayLabel,
    getChartTypeKeyByLabel
} from './chartConfig.js';

const CHART_TYPE_ALIASES = {
    column: 'bar_clustered',
    bar: 'barh_clustered',
    line: 'line_basic',
    pie: 'pie_basic',
    doughnut: 'pie_doughnut',
    area: 'area_basic',
    scatter: 'scatter_basic',
    radar: 'radar_basic',
    histogram: 'histogram_basic'
};

const SUPPORTED_CHART_BASES = new Set([
    'bar',
    'histogram',
    'line',
    'area',
    'pie',
    'scatter',
    'radar',
    'stock',
    'sunburst',
    'treemap',
    'funnel',
    'waterfall',
    'combo'
]);

function toNumber(value, fallback) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
}

function normalizeSheetName(name) {
    const text = String(name ?? '');
    return /[\s']/.test(text) ? `'${text.replaceAll("'", "''")}'` : text;
}

function buildCellRef(sheet, r, c) {
    const sheetName = normalizeSheetName(sheet.name);
    const cellRef = sheet.Utils.cellNumToStr({ r, c });
    return `${sheetName}!${cellRef}`;
}

function buildRangeRef(sheet, sr, sc, er, ec) {
    const sheetName = normalizeSheetName(sheet.name);
    const start = sheet.Utils.cellNumToStr({ r: sr, c: sc });
    const end = sheet.Utils.cellNumToStr({ r: er, c: ec });
    return `${sheetName}!${start}:${end}`;
}

export function resolveChartMeta(SN, chartType) {
    if (!chartType) return null;
    const rawType = String(chartType).trim();
    if (!rawType) return null;

    const lowered = rawType.toLowerCase();
    const localizedKey = getChartTypeKeyByLabel(SN, rawType);
    const normalized = CHART_TYPE_ALIASES[lowered] || localizedKey || rawType;

    let baseKey = normalized;
    let variant = '';
    if (normalized.includes('_')) {
        const parts = normalized.split('_');
        baseKey = parts.shift();
        variant = parts.join('_');
    }

    if (!variant) {
        const defaultVariants = {
            bar: 'clustered',
            barh: 'clustered',
            histogram: 'basic',
            line: 'basic',
            area: 'basic',
            pie: 'basic',
            scatter: 'basic',
            radar: 'basic'
        };
        variant = defaultVariants[baseKey] || '';
    }

    let orientation = 'vertical';
    let base = baseKey;
    if (baseKey === 'barh') {
        base = 'bar';
        orientation = 'horizontal';
    }

    const stacked = variant.includes('stacked') || variant.includes('percent');
    const percent = variant.includes('percent');
    const markers = variant.includes('marker');
    const smooth = variant.includes('smooth');
    const doughnut = variant.includes('doughnut');
    const bubble = variant.includes('bubble');
    const radarFilled = variant.includes('filled');
    const lineLike = variant.includes('line') || smooth;

    const category = CHART_CATEGORIES.find((item) => item.key === baseKey);
    const categoryLabel = category ? SN.t(category.labelKey) : rawType;
    const label = getChartDisplayLabel(SN, normalized, categoryLabel || rawType);

    return {
        key: normalized,
        base,
        baseKey,
        variant,
        label,
        orientation,
        stacked,
        percent,
        markers,
        smooth,
        doughnut,
        radarFilled,
        bubble,
        lineLike,
        supported: SUPPORTED_CHART_BASES.has(base)
    };
}

function buildChartExcelRefs(sheet, area, transposed) {
    if (transposed) {
        const categories = buildRangeRef(sheet, area.s.r + 1, area.s.c, area.e.r, area.s.c);
        const series = [];
        for (let c = area.s.c + 1; c <= area.e.c; c++) {
            series.push({
                name: buildCellRef(sheet, area.s.r, c),
                data: buildRangeRef(sheet, area.s.r + 1, c, area.e.r, c)
            });
        }
        return { categories, series };
    }
    const categories = buildRangeRef(sheet, area.s.r, area.s.c + 1, area.s.r, area.e.c);
    const series = [];
    for (let r = area.s.r + 1; r <= area.e.r; r++) {
        series.push({
            name: buildCellRef(sheet, r, area.s.c),
            data: buildRangeRef(sheet, r, area.s.c + 1, r, area.e.c)
        });
    }
    return { categories, series };
}

function expandRangeRef(ref, SN) {
    if (typeof ref !== 'string' || !ref.includes('!')) return [];
    const match = ref.match(/^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/);
    if (!match) return [];

    const sheetName = (match[1] ?? match[2] ?? '').replace(/''/g, "'");
    const cellRef = match[3].replaceAll('$', '');
    const sheet = SN.getSheet(sheetName);
    if (!sheet) return [];

    const refs = [];
    sheet.eachCells(cellRef, (r, c) => refs.push(buildCellRef(sheet, r, c)));
    return refs;
}

function applyCategoryChartRefs(option, excelRef) {
    if (option.xAxis?.type === 'category') option.xAxis.data = excelRef.categories;
    if (option.yAxis?.type === 'category') option.yAxis.data = excelRef.categories;
    option.series?.forEach((s, i) => {
        if (excelRef.series[i]) {
            s.name = excelRef.series[i].name;
            s.data = excelRef.series[i].data;
        }
    });
    if (option.legend?.data) {
        option.legend.data = excelRef.series.map(s => s.name);
    }
}

function applyChartExcelRefs(SN, option, meta, excelRef) {
    if (!option || !meta || !excelRef) return;

    if (['bar', 'line', 'area', 'combo'].includes(meta.base)) {
        applyCategoryChartRefs(option, excelRef);
        return;
    }

    if (meta.base === 'histogram') {
        if (option.series?.[0] && excelRef.series?.[0]) {
            option.series[0].name = excelRef.series[0].name;
            option.series[0].data = excelRef.series[0].data;
        }
        return;
    }

    if (['pie', 'funnel', 'sunburst', 'treemap'].includes(meta.base)) {
        const categoryRefs = expandRangeRef(excelRef.categories, SN);
        const valueRefs = expandRangeRef(excelRef.series?.[0]?.data, SN);
        const seriesName = excelRef.series?.[0]?.name;

        option.legend = {
            ...(option.legend || {}),
            data: categoryRefs.length ? categoryRefs : option.legend?.data
        };
        option.series?.forEach((s) => {
            s.name = seriesName || s.name;
            s.data = valueRefs.map((value, index) => ({
                ...(s.data?.[index] && typeof s.data[index] === 'object' ? s.data[index] : {}),
                name: categoryRefs[index] ?? s.data?.[index]?.name ?? '',
                value
            }));
        });
        return;
    }

    if (meta.base === 'radar') {
        const categoryRefs = expandRangeRef(excelRef.categories, SN);
        if (categoryRefs.length && option.radar?.indicator) {
            option.radar.indicator = categoryRefs.map((name, index) => ({
                ...(option.radar.indicator[index] || {}),
                name
            }));
        }
        option.series?.forEach((s, i) => {
            if (!excelRef.series[i]) return;
            s.name = excelRef.series[i].name;
            if (s.data?.[0]) {
                s.data[0].name = excelRef.series[i].name;
                s.data[0].value = excelRef.series[i].data;
            }
        });
    }
}

function normalizeSeriesToPercent(series) {
    if (!series.length) return [];
    const length = series[0].data.length;
    const totals = new Array(length).fill(0);

    series.forEach((s) => {
        s.data.forEach((val, idx) => {
            totals[idx] += val;
        });
    });

    return series.map((s) => ({
        name: s.name,
        data: s.data.map((val, idx) => {
            const total = totals[idx];
            if (!total) return 0;
            return Math.round((val / total) * 10000) / 100;
        })
    }));
}

function buildScatterSeries(meta, categories, series) {
    const xValues = categories.map((val, idx) => toNumber(val, idx + 1));
    const seriesType = meta.lineLike ? 'line' : 'scatter';
    const showSymbol = meta.lineLike ? meta.markers : true;

    const buildPairs = (data) => data.map((val, idx) => {
        const xVal = xValues[idx] ?? idx + 1;
        return [xVal, val];
    });

    const buildBubbleData = (data) => data.map((val, idx) => {
        const xVal = xValues[idx] ?? idx + 1;
        if (Array.isArray(val)) {
            const yVal = toNumber(val[1] ?? val[0], 0);
            const sizeVal = toNumber(val.length > 2 ? val[2] : yVal, yVal);
            return [xVal, yVal, sizeVal];
        }
        const yVal = toNumber(val, 0);
        return [xVal, yVal, yVal];
    });

    const buildBubbleSymbolSizer = (data) => {
        const sizes = data.map(item => toNumber(item[2] ?? item[1] ?? item[0], 0));
        const min = Math.min(...sizes);
        const max = Math.max(...sizes);
        const minSize = 6;
        const maxSize = 30;
        if (!Number.isFinite(min) || !Number.isFinite(max) || max === min) {
            const fixed = (minSize + maxSize) / 2;
            return {
                min,
                max,
                symbolSize: () => fixed
            };
        }
        return {
            min,
            max,
            symbolSize: (val) => {
                const raw = Array.isArray(val) ? (val[2] ?? val[1] ?? val[0]) : val;
                const num = toNumber(raw, min);
                return minSize + (maxSize - minSize) * ((num - min) / (max - min));
            }
        };
    };

    return series.map((s) => {
        if (meta.bubble && seriesType === 'scatter') {
            const data = buildBubbleData(s.data);
            const bubbleInfo = buildBubbleSymbolSizer(data);
            return {
                name: s.name,
                type: 'scatter',
                data,
                smooth: false,
                showSymbol: true,
                symbolSize: bubbleInfo.symbolSize,
                __bubble: true,
                __bubbleMin: bubbleInfo.min,
                __bubbleMax: bubbleInfo.max
            };
        }
        return {
            name: s.name,
            type: seriesType,
            data: buildPairs(s.data),
            smooth: meta.smooth,
            showSymbol
        };
    });
}

function extractSeriesValues(seriesData, fallback = 0) {
    if (!Array.isArray(seriesData)) return [];
    return seriesData.map((item, idx) => {
        if (typeof item === 'number') return item;
        if (typeof item === 'string') return toNumber(item, fallback);
        if (Array.isArray(item)) {
            const value = item.length > 1 ? item[1] : item[0];
            return toNumber(value, fallback);
        }
        if (item && typeof item === 'object') {
            if (Object.prototype.hasOwnProperty.call(item, 'value')) {
                return toNumber(item.value, fallback);
            }
        }
        return fallback;
    });
}

function buildNameValueItems(categories, series, seriesIndex = 0, SN = null) {
    const values = extractSeriesValues(series[seriesIndex]?.data);
    return (categories || []).map((name, idx) => ({
        name: String(name ?? (SN ? SN.t('action.insert.content.chart.defaultItem', { index: idx + 1 }) : `Item ${idx + 1}`)),
        value: values[idx] ?? 0
    }));
}

function findSeriesByKeywords(series, keywords) {
    const kw = keywords.map(k => String(k).toLowerCase());
    return series.find((s) => {
        const name = String(s?.name ?? '').toLowerCase();
        return kw.some(k => name.includes(k));
    }) || null;
}

function buildStockOption(SN, meta, categories, series) {
    const openSeries = findSeriesByKeywords(series, ['开盘', 'open']);
    const closeSeries = findSeriesByKeywords(series, ['收盘', 'close']);
    const highSeries = findSeriesByKeywords(series, ['最高', 'high']);
    const lowSeries = findSeriesByKeywords(series, ['最低', 'low']);
    const volumeSeries = findSeriesByKeywords(series, ['成交量', 'volume', 'vol']);

    const remaining = series.filter(s => ![openSeries, closeSeries, highSeries, lowSeries, volumeSeries].includes(s));
    const [fallbackA, fallbackB, fallbackC, fallbackD] = remaining;

    const open = openSeries?.data ?? fallbackA?.data ?? [];
    const high = highSeries?.data ?? fallbackB?.data ?? [];
    const low = lowSeries?.data ?? fallbackC?.data ?? [];
    const close = closeSeries?.data ?? fallbackD?.data ?? open;
    const volume = volumeSeries?.data ?? [];

    const candleData = (categories || []).map((_, idx) => ([
        toNumber(open[idx], 0),
        toNumber(close[idx], 0),
        toNumber(low[idx], 0),
        toNumber(high[idx], 0)
    ]));

    const option = {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { data: [] },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: [{
            name: meta.label,
            type: 'candlestick',
            data: candleData
        }]
    };

    if (volume && volume.length) {
        option.series.push({
            name: SN.t('action.insert.content.chart.volume'),
            type: 'bar',
            data: extractSeriesValues(volume)
        });
    }

    return option;
}

function buildFunnelOption(SN, meta, categories, series) {
    const data = buildNameValueItems(categories, series, 0, SN);
    return {
        title: { text: meta.label },
        tooltip: { trigger: 'item' },
        legend: { show: false },
        series: [{
            name: meta.label,
            type: 'funnel',
            sort: 'descending',
            data
        }]
    };
}

function buildTreeOption(meta, categories, series, type) {
    const data = buildNameValueItems(categories, series, 0);
    const legend = type === 'treemap'
        ? { top: 40, left: 'center', data: categories }
        : { show: false };
    return {
        title: { text: meta.label },
        tooltip: { trigger: 'item' },
        legend,
        series: [{
            name: meta.label,
            type,
            data
        }]
    };
}

function buildWaterfallOption(SN, meta, categories, series) {
    const values = extractSeriesValues(series[0]?.data);
    let cumulative = 0;
    const assist = [];
    const actual = values.map((val) => {
        const start = cumulative;
        cumulative += val;
        assist.push(start);
        return val;
    });
    const seriesName = series[0]?.name || meta.label;
    const colorFn = (params) => {
        const val = actual[params.dataIndex] ?? 0;
        return val >= 0 ? '#70AD47' : '#ED7D31';
    };

    return {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { top: 40, left: 'center', data: [seriesName] },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: [
            {
                name: SN.t('action.insert.content.chart.assist'),
                type: 'bar',
                stack: 'total',
                itemStyle: { color: 'transparent' },
                data: assist
            },
            {
                name: seriesName,
                type: 'bar',
                stack: 'total',
                itemStyle: { color: colorFn },
                data: actual
            }
        ]
    };
}

function buildHistogramOption(meta, categories, series) {
    const values = extractSeriesValues(series[0]?.data);
    const seriesName = series[0]?.name || meta.label;
    return {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { show: false },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: [{
            name: seriesName,
            type: 'bar',
            data: values,
            barGap: '0%',
            barCategoryGap: '0%'
        }]
    };
}

function buildComboOption(meta, categories, series) {
    const option = {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { data: series.map(s => s.name) },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: []
    };

    if (!series.length) return option;
    const barSeries = series[0];
    const lineSeries = series[1];

    if (barSeries) {
        option.series.push({
            name: barSeries.name,
            type: 'bar',
            data: barSeries.data
        });
    }
    if (lineSeries) {
        const lineConfig = {
            name: lineSeries.name,
            type: 'line',
            data: lineSeries.data
        };
        if (meta.variant.includes('area')) {
            lineConfig.areaStyle = {};
        }
        option.series.push(lineConfig);
    } else if (meta.variant.includes('area')) {
        option.series[0].type = 'line';
        option.series[0].areaStyle = {};
    }

    return option;
}

function generateChartOption(SN, meta, categories, series) {
    if (!meta || !meta.supported) return null;

    const title = meta.label || meta.key;
    const option = {
        title: { text: title },
        tooltip: { trigger: meta.base === 'pie' ? 'item' : 'axis' },
        legend: {
            data: meta.base === 'pie' ? categories : series.map(s => s.name)
        }
    };

    if (meta.base === 'stock') {
        return buildStockOption(SN, meta, categories, series);
    }

    if (meta.base === 'sunburst') {
        return buildTreeOption(meta, categories, series, 'sunburst');
    }

    if (meta.base === 'treemap') {
        return buildTreeOption(meta, categories, series, 'treemap');
    }

    if (meta.base === 'funnel') {
        return buildFunnelOption(SN, meta, categories, series);
    }

    if (meta.base === 'waterfall') {
        return buildWaterfallOption(SN, meta, categories, series);
    }

    if (meta.base === 'histogram') {
        return buildHistogramOption(meta, categories, series);
    }

    if (meta.base === 'combo') {
        return buildComboOption(meta, categories, series);
    }

    if (meta.base === 'pie') {
        option.legend.data = categories;
        const pieData = categories.map((name, idx) => ({
            name: name,
            value: series[0]?.data[idx] || 0
        }));

        option.series = [{
            name: series[0]?.name || SN.t('action.insert.content.chart.defaultSeries'),
            type: 'pie',
            radius: meta.doughnut ? ['40%', '70%'] : '60%',
            data: pieData,
            emphasis: {
                itemStyle: {
                    shadowBlur: 10,
                    shadowOffsetX: 0,
                    shadowColor: 'rgba(0, 0, 0, 0.5)'
                }
            },
            label: {
                show: false
            }
        }];
        return option;
    }

    if (meta.base === 'radar') {
        option.radar = {
            indicator: categories.map(name => ({
                name
            }))
        };
        option.series = series.map(s => ({
            name: s.name,
            type: 'radar',
            data: [{
                value: s.data,
                name: s.name
            }],
            symbolSize: meta.markers ? 6 : 0
        }));
        if (meta.radarFilled) {
            option.series = option.series.map(s => ({ ...s, areaStyle: {} }));
        }
        return option;
    }

    if (meta.base === 'scatter') {
        option.xAxis = { type: 'value' };
        option.yAxis = { type: 'value' };
        option.series = buildScatterSeries(meta, categories, series);
        option.tooltip.trigger = meta.lineLike ? 'axis' : 'item';
        return option;
    }

    const usePercent = meta.percent;
    const normalizedSeries = usePercent ? normalizeSeriesToPercent(series) : series;
    const valueAxis = {
        type: 'value',
        max: usePercent ? 100 : undefined,
        axisLabel: usePercent ? { formatter: '{value}%' } : undefined
    };

    if (meta.orientation === 'horizontal') {
        option.xAxis = valueAxis;
        option.yAxis = { type: 'category', data: categories };
    } else {
        option.xAxis = { type: 'category', data: categories };
        option.yAxis = valueAxis;
    }

    option.series = normalizedSeries.map(s => {
        const seriesConfig = {
            name: s.name,
            type: meta.base === 'bar' ? 'bar' : 'line',
            data: s.data
        };

        if (meta.stacked) seriesConfig.stack = 'total';
        if (meta.base === 'area') seriesConfig.areaStyle = {};
        if (meta.base !== 'bar') {
            seriesConfig.showSymbol = meta.markers;
            seriesConfig.smooth = meta.smooth;
        }

        return seriesConfig;
    });

    return option;
}

/**
 * Create an ECharts option from a sheet range.
 * @param {Workbook} SN
 * @param {Sheet} sheet
 * @param {{s: {r: number, c: number}, e: {r: number, c: number}}} area
 * @param {string} chartType
 * @param {Object} [options={}]
 * @param {Object} [options.pivotSource]
 * @returns {{option: Object, meta: Object, excelRef: Object, transposed: boolean}|null}
 */
export function createChartOptionFromRange(SN, sheet, area, chartType, options = {}) {
    if (!SN || !sheet || !area?.s || !area?.e) return null;

    let data = [];
    for (let r = area.s.r; r <= area.e.r; r++) {
        const row = [];
        for (let c = area.s.c; c <= area.e.c; c++) {
            const cell = sheet.getCell(r, c);
            row.push(cell.showVal || '');
        }
        data.push(row);
    }

    if (data.length === 0) return null;

    const dataRows = data.length - 1;
    const dataCols = (data[0]?.length || 1) - 1;
    const transposed = dataRows > dataCols;

    if (transposed) {
        const cols = data[0].length;
        const rows = data.length;
        const newData = [];
        for (let c = 0; c < cols; c++) {
            const row = [];
            for (let r = 0; r < rows; r++) {
                row.push(data[r][c] ?? '');
            }
            newData.push(row);
        }
        data = newData;
    }

    const categories = data[0].slice(1);
    const meta = resolveChartMeta(SN, chartType);
    if (!meta || !meta.supported) return null;

    const series = [];
    for (let i = 1; i < data.length; i++) {
        const seriesName = data[i][0];
        const seriesData = data[i].slice(1).map(v => {
            const num = parseFloat(v);
            return isNaN(num) ? 0 : num;
        });
        series.push({
            name: seriesName,
            data: seriesData
        });
    }

    const option = generateChartOption(SN, meta, categories, series);
    if (!option) return null;
    option.__chartKey = meta.key;
    if (meta.percent) option.__percent = true;

    const excelRef = buildChartExcelRefs(sheet, area, transposed);
    option.__excelRef = excelRef;
    if (options.pivotSource) {
        option.__pivotSource = {
            sheet: options.pivotSource.sheet ?? sheet.name,
            name: options.pivotSource.name ?? null,
            cacheId: options.pivotSource.cacheId ?? null,
            range: options.pivotSource.range ?? null,
            type: options.pivotSource.type ?? meta.key,
            includeTotals: options.pivotSource.includeTotals !== false,
            transposed: options.pivotSource.transposed ?? transposed
        };
    }

    applyChartExcelRefs(SN, option, meta, excelRef);

    if (['sunburst', 'treemap', 'funnel', 'waterfall', 'histogram'].includes(meta.base)) {
        option.__chartType = meta.base;
        option.__excelHidden = excelRef.series.map((_, index) => index > 0);
        if (meta.base === 'waterfall') {
            option.__waterfall = {
                categories: excelRef.categories,
                values: excelRef.series?.[0]?.data || '',
                seriesName: series[0]?.name || meta.label
            };
        } else if (meta.base === 'histogram') {
            option.__histogram = { intervalClosed: 'r' };
        }
    }

    return {
        option,
        meta,
        excelRef,
        transposed
    };
}
