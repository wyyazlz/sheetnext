export const CHART_CATEGORIES = [
    {
        key: 'bar',
        labelKey: 'chart.category.bar',
        items: [
            { key: 'bar_clustered', labelKey: 'chart.item.bar_clustered' },
            { key: 'bar_stacked', labelKey: 'chart.item.bar_stacked' },
            { key: 'bar_percent', labelKey: 'chart.item.bar_percent' }
        ]
    },
    {
        key: 'histogram',
        labelKey: 'chart.category.histogram',
        items: [
            { key: 'histogram_basic', labelKey: 'chart.item.histogram_basic' }
        ]
    },
    {
        key: 'line',
        labelKey: 'chart.category.line',
        items: [
            { key: 'line_basic', labelKey: 'chart.item.line_basic' },
            { key: 'line_stacked', labelKey: 'chart.item.line_stacked' },
            { key: 'line_percent', labelKey: 'chart.item.line_percent' },
            { key: 'line_marker', labelKey: 'chart.item.line_marker' },
            { key: 'line_marker_stacked', labelKey: 'chart.item.line_marker_stacked' },
            { key: 'line_marker_percent', labelKey: 'chart.item.line_marker_percent' }
        ]
    },
    {
        key: 'pie',
        labelKey: 'chart.category.pie',
        items: [
            { key: 'pie_basic', labelKey: 'chart.item.pie_basic' },
            { key: 'pie_doughnut', labelKey: 'chart.item.pie_doughnut' }
        ]
    },
    {
        key: 'barh',
        labelKey: 'chart.category.barh',
        items: [
            { key: 'barh_clustered', labelKey: 'chart.item.barh_clustered' },
            { key: 'barh_stacked', labelKey: 'chart.item.barh_stacked' },
            { key: 'barh_percent', labelKey: 'chart.item.barh_percent' }
        ]
    },
    {
        key: 'area',
        labelKey: 'chart.category.area',
        items: [
            { key: 'area_basic', labelKey: 'chart.item.area_basic' },
            { key: 'area_stacked', labelKey: 'chart.item.area_stacked' },
            { key: 'area_percent', labelKey: 'chart.item.area_percent' }
        ]
    },
    {
        key: 'scatter',
        labelKey: 'chart.category.scatter',
        items: [
            { key: 'scatter_basic', labelKey: 'chart.item.scatter_basic' },
            { key: 'scatter_smooth_marker', labelKey: 'chart.item.scatter_smooth_marker' },
            { key: 'scatter_smooth', labelKey: 'chart.item.scatter_smooth' },
            { key: 'scatter_line_marker', labelKey: 'chart.item.scatter_line_marker' },
            { key: 'scatter_line', labelKey: 'chart.item.scatter_line' },
            { key: 'scatter_bubble', labelKey: 'chart.item.scatter_bubble' }
        ]
    },
    {
        key: 'stock',
        labelKey: 'chart.category.stock',
        items: [
            { key: 'stock_hlc', labelKey: 'chart.item.stock_hlc' },
            { key: 'stock_ohlc', labelKey: 'chart.item.stock_ohlc' },
            { key: 'stock_vol_hlc', labelKey: 'chart.item.stock_vol_hlc' },
            { key: 'stock_vol_ohlc', labelKey: 'chart.item.stock_vol_ohlc' }
        ]
    },
    {
        key: 'radar',
        labelKey: 'chart.category.radar',
        items: [
            { key: 'radar_basic', labelKey: 'chart.item.radar_basic' },
            { key: 'radar_marker', labelKey: 'chart.item.radar_marker' },
            { key: 'radar_filled', labelKey: 'chart.item.radar_filled' }
        ]
    },
    {
        key: 'sunburst',
        labelKey: 'chart.category.sunburst',
        items: [
            { key: 'sunburst_basic', labelKey: 'chart.item.sunburst_basic' }
        ]
    },
    {
        key: 'treemap',
        labelKey: 'chart.category.treemap',
        items: [
            { key: 'treemap_basic', labelKey: 'chart.item.treemap_basic' }
        ]
    },
    {
        key: 'funnel',
        labelKey: 'chart.category.funnel',
        items: [
            { key: 'funnel_basic', labelKey: 'chart.item.funnel_basic' }
        ]
    },
    {
        key: 'waterfall',
        labelKey: 'chart.category.waterfall',
        items: [
            { key: 'waterfall_basic', labelKey: 'chart.item.waterfall_basic' }
        ]
    },
    {
        key: 'combo',
        labelKey: 'chart.category.combo',
        items: [
            { key: 'combo_bar_line', labelKey: 'chart.item.combo_bar_line' },
            { key: 'combo_area_bar', labelKey: 'chart.item.combo_area_bar' }
        ]
    }
];

export const CHART_SHORTCUT_MENUS = {
    chartBar: 'bar',
    chartLine: 'line',
    chartPie: 'pie',
    chartArea: 'area',
    chartScatter: 'scatter',
    chartRadar: 'radar'
};

export function getLocalizedChartCategories(SN) {
    return CHART_CATEGORIES.map((category) => ({
        ...category,
        text: SN.t(category.labelKey),
        items: category.items.map((item) => ({
            ...item,
            text: SN.t(item.labelKey)
        }))
    }));
}

export function getChartItemByKey(itemKey) {
    for (const category of CHART_CATEGORIES) {
        const item = category.items.find((chartItem) => chartItem.key === itemKey);
        if (item) return { ...item, category };
    }
    return null;
}

export function getChartLabelKeyByItemKey(itemKey) {
    const item = getChartItemByKey(itemKey);
    return item?.labelKey || '';
}

export function getChartDisplayLabel(SN, itemKey, fallback = '') {
    const item = getChartItemByKey(itemKey);
    if (!item) return fallback || itemKey || '';
    return SN.t(item.labelKey);
}

export function getChartTypeKeyByLabel(SN, chartTypeLabel) {
    const label = typeof chartTypeLabel === 'string' ? chartTypeLabel.trim() : '';
    if (!label) return '';

    for (const category of CHART_CATEGORIES) {
        if (category.key === label || SN.t(category.labelKey) === label) return category.key;
        for (const item of category.items) {
            if (item.key === label || SN.t(item.labelKey) === label) return item.key;
        }
    }
    return '';
}
