
import { numFmtFun } from "../Cell/helpers.js";
import { getColor } from "../Xml/helpers.js";

export const chartXmlToEChartOption = (obj, SN) => {
    if (!obj) return null;
    if (obj['cx:chartSpace'] || obj['cx:chartData']) {
        return chartExXmlToEChartOption(obj, SN);
    }
    obj = obj['c:chartSpace']
    const chart = obj['c:chart'];
    const plotArea = chart['c:plotArea'];
    const legendXml = chart['c:legend'];
    const titleXml = chart['c:title'];
    const isPivotChart = !!obj['c:pivotSource'];

    const baseOption = {
        title: getTitle(titleXml, SN),
        legend: getLegend(legendXml, SN),
        series: [],
        ...(isPivotChart ? { __pivotChart: true } : {})
    };
    const chartAreaStyle = getShapeStyleOption(obj?.['c:spPr'], SN);
    const plotAreaStyle = getShapeStyleOption(plotArea?.['c:spPr'], SN);
    if (chartAreaStyle.fillColor) {
        baseOption.backgroundColor = chartAreaStyle.fillColor;
    }
    if (chartAreaStyle.lineStyle) {
        baseOption.__excelChartAreaBorder = chartAreaStyle.lineStyle;
    }
    if (obj?.['c:spPr']) {
        baseOption.__excelChartAreaSpPr = cloneChartXmlNode(obj['c:spPr']);
    }
    if (obj?.['c:txPr']) {
        baseOption.__excelChartSpaceTxPr = cloneChartXmlNode(obj['c:txPr']);
    }
    if (plotArea?.['c:spPr']) {
        baseOption.__excelPlotAreaSpPr = cloneChartXmlNode(plotArea['c:spPr']);
    }
    if (plotAreaStyle.lineStyle) {
        baseOption.__excelPlotAreaBorder = plotAreaStyle.lineStyle;
    }
    if (plotAreaStyle.fillColor || plotAreaStyle.lineStyle?.show === true) {
        baseOption.grid = {
            ...(baseOption.grid || {}),
            show: true,
            ...(plotAreaStyle.fillColor ? { backgroundColor: plotAreaStyle.fillColor } : {}),
            ...(plotAreaStyle.lineStyle?.color ? { borderColor: plotAreaStyle.lineStyle.color } : {}),
            ...(plotAreaStyle.lineStyle?.width != null ? { borderWidth: plotAreaStyle.lineStyle.width } : {}),
            ...(plotAreaStyle.lineStyle?.type ? { borderType: plotAreaStyle.lineStyle.type } : {})
        };
    }
    const chartSpaceTextStyle = getTextStyleFromTxPr(obj?.['c:txPr'], SN, 13.33);
    if (Object.keys(chartSpaceTextStyle).length > 0) {
        baseOption.__excelChartSpaceTextStyle = chartSpaceTextStyle;
    }

    const buildSeries = (ser, type, index, extra = {}) => {
        let dataRefOrData;

        switch (type) {
            case 'scatter':
            case 'bubble': {
                let xArr = []
                const xRef = safeGet(ser, ['c:xVal', 'c:numRef', 'c:f']) || safeGet(ser, ['c:xVal', 'c:strRef', 'c:f']);
                if (xRef) {
                    xArr = rangeStrToCellRefs(xRef, SN)
                    if (isPivotChart && xArr.length === 0) {
                        xArr = getCachedPoints(safeGet(ser, ['c:xVal', 'c:numRef']) || safeGet(ser, ['c:xVal', 'c:strRef']));
                    }
                }
                let yArr = []
                const yRef = safeGet(ser, ['c:yVal', 'c:numRef', 'c:f']) || safeGet(ser, ['c:yVal', 'c:strRef', 'c:f']);
                if (yRef) {
                    yArr = rangeStrToCellRefs(yRef, SN)
                    if (isPivotChart && yArr.length === 0) {
                        yArr = getCachedPoints(safeGet(ser, ['c:yVal', 'c:numRef']) || safeGet(ser, ['c:yVal', 'c:strRef']));
                    }
                }
                if (type === 'bubble') {
                    let sizeArr = []
                    const sizeRef = safeGet(ser, ['c:bubbleSize', 'c:numRef', 'c:f']);
                    if (sizeRef) {
                        sizeArr = rangeStrToCellRefs(sizeRef, SN)
                        if (isPivotChart && sizeArr.length === 0) {
                            sizeArr = getCachedPoints(safeGet(ser, ['c:bubbleSize', 'c:numRef']));
                        }
                    }
                    dataRefOrData = xArr.map((x, i) => [x, yArr[i], sizeArr[i]]);
                } else {
                    dataRefOrData = xArr.map((x, i) => [x, yArr[i]]);
                }
                break;
            }

            case 'pie': {
                let values = []
                let names = []
                const valueRefNode = safeGet(ser, ['c:val', 'c:numRef']);
                const valueRef = valueRefNode?.['c:f'];
                const valueCache = isPivotChart ? getCachedPoints(valueRefNode) : [];
                if (valueRef) {
                    values = rangeStrToCellRefs(valueRef, SN)
                } else if (valueCache.length) {
                    values = valueCache;
                } else {
                    const valArr = safeGet(ser, ['c:val', 'c:numLit', 'c:pt']);
                    if (valArr) values = valArr.map(item => item['c:v'])
                }

                const nameRefNode = safeGet(ser, ['c:cat', 'c:strRef']) ||
                    safeGet(ser, ['c:cat', 'c:numRef']) ||
                    safeGet(ser, ['c:cat', 'c:multiLvlStrRef']);
                const nameRef = nameRefNode?.['c:f'];
                const nameCache = isPivotChart ? getCachedCategoryPoints(nameRefNode) : [];
                if (nameRef) {
                    names = rangeStrToCellRefs(nameRef, SN)
                    if (isPivotChart && names.length === 0) names = nameCache;
                } else if (nameCache.length) {
                    names = nameCache;
                } else {
                    const nameArr = safeGet(ser, ['c:cat', 'c:strLit', 'c:pt']) || safeGet(ser, ['c:cat', 'c:numLit', 'c:pt']);
                    if (nameArr) names = nameArr.map(item => item['c:v'])
                }
                const pointMap = new Map(
                    sureArray(ser?.['c:dPt']).map((point) => [
                        Number(point?.['c:idx']?._$val),
                        getPointItemStyle(point, SN)
                    ])
                );
                dataRefOrData = values.map((item, dataIndex) => {
                    const pointStyle = pointMap.get(dataIndex);
                    return {
                        name: names[dataIndex],
                        value: item,
                        ...(pointStyle ? { itemStyle: pointStyle } : {})
                    };
                })
                break;
            }

            case 'radar': {
                const valRefNode = safeGet(ser, ['c:val', 'c:numRef']);
                const valRef = valRefNode?.['c:f'];
                const valCache = isPivotChart ? getCachedPoints(valRefNode) : [];
                if (valRef) {
                    dataRefOrData = [{ value: valRef }];
                    break;
                }
                if (valCache.length) {
                    dataRefOrData = [{ value: valCache }];
                    break;
                }
                const valArr = safeGet(ser, ['c:val', 'c:numLit', 'c:pt']);
                dataRefOrData = valArr ? valArr.map(pt => pt['c:v']) : [];
                break;
            }

            default: {
                const valRefNode = safeGet(ser, ['c:val', 'c:numRef']);
                const valRef = valRefNode?.['c:f'];
                const valCache = isPivotChart ? getCachedPoints(valRefNode) : [];
                if (valRef) {
                    dataRefOrData = valRef;
                    break;
                }
                if (valCache.length) {
                    dataRefOrData = valCache;
                    break;
                }
                const valArr = safeGet(ser, ['c:val', 'c:numLit', 'c:pt']);
                dataRefOrData = valArr ? valArr.map(pt => pt['c:v']) : [];
            }
        }

        return {
            name: getSeriesName(ser, index, isPivotChart),
            type: type === 'bubble' ? 'scatter' : type,
            data: dataRefOrData,
            ...extra
        };
    };

    const buildSeriesStyleExtra = (ser, type, extra = {}, workbook = null) => {
        const result = { ...extra };
        const strokeColor = getSeriesStrokeColor(ser, workbook);
        const fillColor = getSeriesFillColor(ser, workbook);
        const lineWidth = getSeriesLineWidth(ser);
        const dashType = getLineDashType(ser);
        if (type === 'line' || type === 'scatter' || type === 'radar') {
            if (strokeColor || lineWidth != null || dashType) {
                result.lineStyle = {
                    ...(result.lineStyle || {}),
                    ...(strokeColor ? { color: strokeColor } : {}),
                    ...(lineWidth != null ? { width: lineWidth } : {}),
                    ...(dashType ? { type: dashType } : {})
                };
            }
        }
        if (fillColor && result.areaStyle != null) {
            result.areaStyle = {
                ...(typeof result.areaStyle === 'object' ? result.areaStyle : {}),
                color: fillColor
            };
        } else if ((type === 'scatter' || type === 'bar' || type === 'radar') && (fillColor || strokeColor || lineWidth != null)) {
            result.itemStyle = {
                ...(result.itemStyle || {}),
                ...(fillColor ? { color: fillColor } : {}),
                ...(strokeColor ? { borderColor: strokeColor } : {}),
                ...(lineWidth != null ? { borderWidth: lineWidth } : {}),
                ...(dashType ? { borderType: dashType } : {})
            };
        }
        return result;
    };

    const hasScatterLike = !!plotArea['c:scatterChart'] || !!plotArea['c:bubbleChart'];
    const hasCategoryLike = !!plotArea['c:lineChart'] || !!plotArea['c:line3DChart'] ||
        !!plotArea['c:barChart'] || !!plotArea['c:bar3DChart'] ||
        !!plotArea['c:areaChart'] || !!plotArea['c:area3DChart'] ||
        !!plotArea['c:radarChart'];

    if (hasScatterLike && !hasCategoryLike) {
        const valAxes = sureArray(plotArea['c:valAx']);
        const xAxisOptions = [];
        const yAxisOptions = [];
        const xAxisIdMap = new Map();
        const yAxisIdMap = new Map();

        valAxes.forEach((axisObj) => {
            const axisId = getAxisId(axisObj);
            const pos = axisObj?.['c:axPos']?._$val;
            if (pos === 't' || pos === 'b') {
                const axisIndex = xAxisOptions.length;
                xAxisOptions.push(buildValueAxisOption(axisObj, SN, pos === 't' ? 'top' : 'bottom'));
                xAxisIdMap.set(axisId, axisIndex);
            } else {
                const axisIndex = yAxisOptions.length;
                yAxisOptions.push(buildValueAxisOption(axisObj, SN, pos === 'r' ? 'right' : 'left'));
                yAxisIdMap.set(axisId, axisIndex);
            }
        });

        sureArray(plotArea['c:scatterChart']).forEach((chartGroup) => {
            const axisIds = sureArray(chartGroup?.['c:axId']).map(item => String(item?._$val ?? ''));
            const xAxisIndex = xAxisIdMap.get(axisIds.find(id => xAxisIdMap.has(id))) ?? 0;
            const yAxisIndex = yAxisIdMap.get(axisIds.find(id => yAxisIdMap.has(id))) ?? 0;
            const arr = sureArray(chartGroup?.['c:ser']);
            const groupLabelNode = chartGroup?.['c:dLbls'];
            arr.forEach((ser) => {
                const marker = ser?.['c:marker'];
                const symbol = marker?.['c:symbol']?._$val;
                const size = Number(marker?.['c:size']?._$val);
                const showSymbol = symbol ? symbol !== 'none' : true;
                const symbolSize = Number.isFinite(size) && size > 0 ? (size * 96 / 72) : undefined;
                const smooth = ser?.['c:smooth']?._$val === '1';
                const symbolType = getMarkerSymbol(marker);
                const labelExtra = buildSeriesLabelOption(mergeLabelNodes(groupLabelNode, ser?.['c:dLbls']), 'scatter', SN);
                const seriesObj = buildSeries(
                    ser,
                    'scatter',
                    baseOption.series.length,
                    buildSeriesStyleExtra(ser, 'scatter', {
                        showSymbol,
                        ...(symbolType ? { symbol: symbolType } : {}),
                        ...(symbolSize ? { symbolSize } : {}),
                        ...(smooth ? { smooth: true } : {}),
                        ...labelExtra
                    }, SN)
                );
                baseOption.series.push(attachAxisIndices(seriesObj, xAxisIndex, yAxisIndex));
            });
        });

        sureArray(plotArea['c:bubbleChart']).forEach((chartGroup) => {
            const axisIds = sureArray(chartGroup?.['c:axId']).map(item => String(item?._$val ?? ''));
            const xAxisIndex = xAxisIdMap.get(axisIds.find(id => xAxisIdMap.has(id))) ?? 0;
            const yAxisIndex = yAxisIdMap.get(axisIds.find(id => yAxisIdMap.has(id))) ?? 0;
            const arr = sureArray(chartGroup?.['c:ser']);
            const groupLabelNode = chartGroup?.['c:dLbls'];
            arr.forEach((ser) => {
                const seriesObj = buildSeries(
                    ser,
                    'bubble',
                    baseOption.series.length,
                    buildSeriesStyleExtra(ser, 'scatter', {
                        __bubble: true,
                        ...buildSeriesLabelOption(mergeLabelNodes(groupLabelNode, ser?.['c:dLbls']), 'scatter', SN)
                    }, SN)
                );
                baseOption.series.push(attachAxisIndices(seriesObj, xAxisIndex, yAxisIndex));
            });
        });

        baseOption.xAxis = xAxisOptions.length <= 1 ? (xAxisOptions[0] || { type: 'value' }) : xAxisOptions;
        baseOption.yAxis = yAxisOptions.length <= 1 ? (yAxisOptions[0] || { type: 'value' }) : yAxisOptions;
    } else {
        const catAxes = sureArray(plotArea['c:catAx']);
        const valAxes = sureArray(plotArea['c:valAx']);
        const catAxisOptions = catAxes.map(axisObj => buildCategoryAxisOption(axisObj, SN));
        const valAxisOptions = valAxes.map(axisObj => buildValueAxisOption(axisObj, SN));
        const catAxisIdMap = new Map(catAxes.map((axisObj, axisIndex) => [getAxisId(axisObj), axisIndex]));
        const valAxisIdMap = new Map(valAxes.map((axisObj, axisIndex) => [getAxisId(axisObj), axisIndex]));

        const applyCategoryData = (axisIndex, data) => {
            if (axisIndex < 0 || axisIndex >= catAxisOptions.length) return;
            const axis = catAxisOptions[axisIndex];
            if (axis.data == null && data !== undefined) {
                axis.data = data;
            }
        };

        const orderedCategoryGroups = [];
        Object.keys(plotArea).forEach((key) => {
            if (![
                'c:lineChart',
                'c:line3DChart',
                'c:areaChart',
                'c:area3DChart',
                'c:barChart',
                'c:bar3DChart',
                'c:radarChart'
            ].includes(key)) return;
            sureArray(plotArea[key]).forEach((chartGroup) => {
                orderedCategoryGroups.push({ key, chartGroup });
            });
        });

        orderedCategoryGroups.forEach(({ key, chartGroup }) => {
            if (key === 'c:radarChart') {
                const arr = sureArray(chartGroup?.['c:ser']);
                const groupLabelNode = chartGroup?.['c:dLbls'];
                const axisIds = sureArray(chartGroup?.['c:axId']).map(item => String(item?._$val ?? ''));
                const categoryAxisObj = catAxes[catAxisIdMap.get(axisIds.find(id => catAxisIdMap.has(id))) ?? 0];
                const valueAxisObj = valAxes[valAxisIdMap.get(axisIds.find(id => valAxisIdMap.has(id))) ?? 0];
                let qz = getCategoryData(arr[0], isPivotChart);
                if (typeof qz == 'string') {
                    qz = rangeStrToCellRefs(qz, SN);
                } else if (Array.isArray(qz) && qz[0]?.['c:v'] !== undefined) {
                    qz = qz.map(pt => pt['c:v']);
                }
                baseOption.radar = buildRadarCoordinateOption(qz, categoryAxisObj, valueAxisObj, SN);
                baseOption.series[0] = {
                    name: '',
                    type: 'radar',
                    ...buildSeriesLabelOption(groupLabelNode, 'radar', SN),
                    data: arr.map((ser, dataIndex) => {
                        const data = buildSeries(ser, 'radar', dataIndex);
                        const item = {
                            name: data.name,
                            value: data.data[0].value,
                            ...buildSeriesStyleExtra(ser, 'radar', {}, SN)
                        };
                        const mergedLabelNode = mergeLabelNodes(groupLabelNode, ser?.['c:dLbls']);
                        const itemLabelOption = buildSeriesLabelOption(mergedLabelNode, 'radar', SN);
                        if (itemLabelOption.label) {
                            item.label = itemLabelOption.label;
                        } else if (ser?.['c:dLbls']) {
                            item.label = { show: false };
                        }
                        return item;
                    })
                };
                return;
            }

            const axisIds = sureArray(chartGroup?.['c:axId']).map(item => String(item?._$val ?? ''));
            const xAxisIndex = catAxisIdMap.get(axisIds.find(id => catAxisIdMap.has(id))) ?? 0;
            const yAxisIndex = valAxisIdMap.get(axisIds.find(id => valAxisIdMap.has(id))) ?? 0;
            const arr = sureArray(chartGroup?.['c:ser']);
            const grouping = chartGroup?.['c:grouping']?._$val;
            const groupLabelNode = chartGroup?.['c:dLbls'];
            if (grouping === 'percentStacked') baseOption.__percent = true;

            if (key === 'c:barChart' || key === 'c:bar3DChart') {
                const dir = chartGroup?.['c:barDir']?._$val;
                const isHorizontal = dir === 'bar';
                applyCategoryData(xAxisIndex, extractCategories(arr[0], isPivotChart));
                arr.forEach((ser) => {
                    const seriesObj = buildSeries(
                        ser,
                        'bar',
                        baseOption.series.length,
                        buildSeriesStyleExtra(ser, 'bar', {
                            ...(grouping === 'stacked' || grouping === 'percentStacked' ? { stack: 'total' } : {}),
                            ...buildSeriesLabelOption(mergeLabelNodes(groupLabelNode, ser?.['c:dLbls']), 'bar', SN)
                        }, SN)
                    );
                    baseOption.series.push(attachAxisIndices(
                        seriesObj,
                        isHorizontal ? yAxisIndex : xAxisIndex,
                        isHorizontal ? xAxisIndex : yAxisIndex
                    ));
                });
                return;
            }

            applyCategoryData(xAxisIndex, extractCategories(arr[0], isPivotChart));
            arr.forEach((ser) => {
                const marker = ser?.['c:marker'];
                const symbol = marker?.['c:symbol']?._$val;
                const size = Number(marker?.['c:size']?._$val);
                const showSymbol = symbol ? symbol !== 'none' : false;
                const symbolSize = Number.isFinite(size) && size > 0 ? (size * 96 / 72) : undefined;
                const smooth = ser?.['c:smooth']?._$val === '1';
                const symbolType = getMarkerSymbol(marker);
                const seriesObj = buildSeries(
                    ser,
                    'line',
                    baseOption.series.length,
                    buildSeriesStyleExtra(ser, 'line', {
                        ...(key === 'c:areaChart' || key === 'c:area3DChart' ? { areaStyle: {} } : {}),
                        ...(grouping === 'stacked' || grouping === 'percentStacked' ? { stack: 'total' } : {}),
                        showSymbol,
                        ...(symbolType ? { symbol: symbolType } : {}),
                        ...(symbolSize ? { symbolSize } : {}),
                        ...(smooth ? { smooth: true } : {}),
                        ...buildSeriesLabelOption(mergeLabelNodes(groupLabelNode, ser?.['c:dLbls']), 'line', SN)
                    }, SN)
                );
                baseOption.series.push(attachAxisIndices(seriesObj, xAxisIndex, yAxisIndex));
            });
        });

        if (plotArea['c:pieChart'] || plotArea['c:pie3DChart']) {
            const chartNode = plotArea['c:pieChart'] || plotArea['c:pie3DChart'];
            const ser = sureArray(chartNode['c:ser'])[0];
            if (ser) {
                const startAngle = Number(chartNode?.['c:firstSliceAng']?._$val);
                const pieSeries = buildSeries(ser, 'pie', baseOption.series.length, {
                    ...(Number.isFinite(startAngle) ? { startAngle } : {}),
                    ...buildSeriesLabelOption(mergeLabelNodes(chartNode?.['c:dLbls'], ser?.['c:dLbls']), 'pie', SN)
                });
                baseOption.series.push(pieSeries);
            }
        }
        if (plotArea['c:doughnutChart']) {
            const chartNode = plotArea['c:doughnutChart'];
            const ser = sureArray(chartNode['c:ser'])[0];
            if (ser) {
                const startAngle = Number(chartNode?.['c:firstSliceAng']?._$val);
                const holeSize = Number(chartNode?.['c:holeSize']?._$val);
                const outerRadius = 50;
                const innerRadius = Number.isFinite(holeSize) ? Math.round((outerRadius * holeSize) / 100) : 40;
                const doughnutSeries = buildSeries(
                    ser,
                    'pie',
                    baseOption.series.length,
                    {
                        radius: [`${innerRadius}%`, `${outerRadius}%`],
                        ...(Number.isFinite(startAngle) ? { startAngle } : {}),
                        ...buildSeriesLabelOption(mergeLabelNodes(chartNode?.['c:dLbls'], ser?.['c:dLbls']), 'pie', SN)
                    }
                );
                baseOption.series.push(doughnutSeries);
            }
        }

        const barGroups = [
            ...sureArray(plotArea['c:barChart']),
            ...sureArray(plotArea['c:bar3DChart'])
        ];
        const hasHorizontalBar = barGroups.some(chartGroup => chartGroup?.['c:barDir']?._$val === 'bar');
        const hasVerticalCategoryChart = !!plotArea['c:lineChart'] || !!plotArea['c:line3DChart'] ||
            !!plotArea['c:areaChart'] || !!plotArea['c:area3DChart'] || !!plotArea['c:radarChart'] ||
            barGroups.some(chartGroup => chartGroup?.['c:barDir']?._$val !== 'bar');
        const hasCartesianCategoryChart = !!plotArea['c:lineChart'] || !!plotArea['c:line3DChart'] ||
            !!plotArea['c:areaChart'] || !!plotArea['c:area3DChart'] ||
            barGroups.some(chartGroup => chartGroup?.['c:barDir']?._$val !== 'bar');
        const isRadarOnlyChart = !!plotArea['c:radarChart'] && !hasHorizontalBar && !hasCartesianCategoryChart;

        if ((catAxisOptions.length > 0 || valAxisOptions.length > 0) && !isRadarOnlyChart) {
            if (hasHorizontalBar && !hasVerticalCategoryChart) {
                baseOption.xAxis = valAxisOptions.length <= 1 ? (valAxisOptions[0] || { type: 'value' }) : valAxisOptions;
                baseOption.yAxis = catAxisOptions.length <= 1 ? (catAxisOptions[0] || { type: 'category' }) : catAxisOptions;
            } else {
                baseOption.xAxis = catAxisOptions.length <= 1 ? (catAxisOptions[0] || { type: 'category' }) : catAxisOptions;
                baseOption.yAxis = valAxisOptions.length <= 1 ? (valAxisOptions[0] || { type: 'value' }) : valAxisOptions;
            }
            if (!Array.isArray(baseOption.xAxis) && baseOption.xAxis?.type === 'category' && Array.isArray(baseOption.xAxis.data) && baseOption.xAxis.data.length === 0) {
                delete baseOption.xAxis.data;
            }
            if (!Array.isArray(baseOption.yAxis) && baseOption.yAxis?.type === 'category' && Array.isArray(baseOption.yAxis.data) && baseOption.yAxis.data.length === 0) {
                delete baseOption.yAxis.data;
            }
        }
    }

    if (isPivotChart) {
        applyPivotChartPresentation(baseOption, obj, plotArea, SN);
    }

    const ml = safeGet(plotArea, ['c:layout', 'c:manualLayout']);
    if (ml) {
        const x = parseFloat(ml['c:x']?._$val);
        const y = parseFloat(ml['c:y']?._$val);
        const w = parseFloat(ml['c:w']?._$val);
        const h = parseFloat(ml['c:h']?._$val);
        if (!isNaN(x) && !isNaN(y) && !isNaN(w) && !isNaN(h)) {
            baseOption.grid = {
                ...(baseOption.grid || {}),
                left: `${x * 100}%`,
                top: `${y * 100}%`,
                right: `${(1 - x - w) * 100}%`,
                bottom: `${(1 - y - h) * 100}%`,
                containLabel: true
            };
        }
    }

    return chartOptionApplyTemplate(baseOption, SN);

};

function getChartExText(txObj) {
    const v = txObj?.['cx:txData']?.['cx:v'];
    if (typeof v === 'string') return v;
    if (Array.isArray(v)) return v.join('');
    return '';
}

function getChartExTitle(titleObj) {
    const text = getChartExText(titleObj?.['cx:tx']);
    return { text, show: text !== '' };
}

function getChartExLegend(legendObj) {
    if (!legendObj) return { show: false };
    const pos = legendObj?._$pos;
    const legendOption = { orient: 'horizontal' };
    switch (pos) {
        case 't':
            legendOption.top = 40;
            legendOption.left = 'center';
            break;
        case 'l':
            legendOption.left = 'left';
            legendOption.top = 'middle';
            legendOption.orient = 'vertical';
            break;
        case 'r':
            legendOption.right = 'right';
            legendOption.top = 'middle';
            legendOption.orient = 'vertical';
            break;
        case 'b':
            legendOption.bottom = 10;
            legendOption.left = 'center';
            break;
        default:
            legendOption.bottom = 10;
            legendOption.left = 'center';
    }
    return legendOption;
}

function getChartExDimRef(dimObj, SN) {
    const f = dimObj?.['cx:f'];
    const raw = typeof f === 'string' ? f : f?.['#text'];
    if (!raw) return '';
    return resolveDefinedNameRef(raw, SN);
}

function getChartExSeriesNameRef(seriesObj, SN) {
    const f = seriesObj?.['cx:tx']?.['cx:txData']?.['cx:f'];
    if (typeof f === 'string') return resolveDefinedNameRef(f, SN);
    const v = seriesObj?.['cx:tx']?.['cx:txData']?.['cx:v'];
    return typeof v === 'string' ? v : '';
}

function buildChartExCategoryValuePairs(catRefs, valRefs) {
    const maxLen = Math.max(catRefs.length, valRefs.length);
    const data = [];
    for (let i = 0; i < maxLen; i++) {
        data.push({
            name: catRefs[i] ?? `\u7c7b\u522b${i + 1}`,
            value: valRefs[i] ?? 0
        });
    }
    return data;
}

function chartExXmlToEChartOption(obj, SN) {
    const chartSpace = obj['cx:chartSpace'] || obj;
    const chart = chartSpace?.['cx:chart'];
    const plotArea = chart?.['cx:plotArea'];
    const plotRegion = plotArea?.['cx:plotAreaRegion'];
    const seriesArr = sureArray(plotRegion?.['cx:series']);
    const dataArr = sureArray(chartSpace?.['cx:chartData']?.['cx:data']);
    const dataMap = new Map();

    dataArr.forEach((d) => {
        const id = d?._$id;
        if (id == null) return;
        const catRef = getChartExDimRef(d?.['cx:strDim'], SN);
        const valRef = getChartExDimRef(d?.['cx:numDim'], SN);
        dataMap.set(String(id), { catRef, valRef });
    });

    const layoutId = seriesArr[0]?._$layoutId;
    const hiddenFlags = seriesArr.map(s => s?._$hidden === '1');
    const hasBinning = seriesArr.some(s => s?.['cx:layoutPr']?.['cx:binning']);
    const option = {
        title: getChartExTitle(chart?.['cx:title']),
        legend: getChartExLegend(chart?.['cx:legend']),
        series: []
    };

    if (!layoutId) return chartOptionApplyTemplate(option, SN);

    const visibleSeries = seriesArr.filter(s => s?._$hidden !== '1');
    const activeSeries = visibleSeries.length ? visibleSeries : seriesArr;
    const baseSeries = activeSeries[0] || seriesArr[0];
    const baseDataId = String(baseSeries?.['cx:dataId']?._$val ?? 0);
    const baseRefs = dataMap.get(baseDataId) || {};
    const excelSeries = seriesArr.map((ser, index) => {
        const dataId = String(ser?.['cx:dataId']?._$val ?? index);
        const refs = dataMap.get(dataId) || {};
        return {
            name: getChartExSeriesNameRef(ser, SN) || '\u7cfb\u5217' + (index + 1),
            data: refs.valRef || ''
        };
    });

    if (layoutId === 'sunburst' || layoutId === 'treemap') {
        const type = layoutId;
        if (activeSeries.length === 1) {
            const catRefs = rangeStrToCellRefs(baseRefs.catRef, SN);
            const valRefs = rangeStrToCellRefs(baseRefs.valRef, SN);
            const data = buildChartExCategoryValuePairs(catRefs, valRefs);
            option.series = [{
                name: getChartExSeriesNameRef(baseSeries, SN) || '',
                type,
                data
            }];
            if (type === 'treemap') {
                option.legend = { ...(option.legend || {}), show: true, data: data.map(item => item.name), top: 40, left: 'center' };
            } else {
                option.legend = { ...(option.legend || {}), show: false };
            }
        } else {
            const nodes = activeSeries.map((ser, index) => {
                const dataId = String(ser?.['cx:dataId']?._$val ?? index);
                const refs = dataMap.get(dataId) || {};
                const catRefs = rangeStrToCellRefs(refs.catRef, SN);
                const valRefs = rangeStrToCellRefs(refs.valRef, SN);
                const children = buildChartExCategoryValuePairs(catRefs, valRefs);
                return {
                    name: getChartExSeriesNameRef(ser, SN) || '\u7cfb\u5217' + (index + 1),
                    children
                };
            });
            option.series = [{
                type,
                data: nodes
            }];
        }
        option.__chartType = type;
        option.__excelHidden = hiddenFlags;
        option.__excelRef = {
            categories: baseRefs.catRef || '',
            series: excelSeries
        };
    } else if (layoutId === 'funnel') {
        const catRefs = rangeStrToCellRefs(baseRefs.catRef, SN);
        const valRefs = rangeStrToCellRefs(baseRefs.valRef, SN);
        const data = buildChartExCategoryValuePairs(catRefs, valRefs);
        option.series = [{
            name: getChartExSeriesNameRef(baseSeries, SN) || '',
            type: 'funnel',
            data
        }];
        option.__chartType = 'funnel';
        option.__excelHidden = hiddenFlags;
        option.__excelRef = {
            categories: baseRefs.catRef || '',
            series: excelSeries
        };
    } else if (layoutId === 'waterfall') {
        const catRefs = rangeStrToCellRefs(baseRefs.catRef, SN);
        const valRefs = rangeStrToCellRefs(baseRefs.valRef, SN);
        option.__chartType = 'waterfall';
        option.__waterfall = {
            categories: catRefs,
            values: valRefs,
            seriesName: getChartExSeriesNameRef(baseSeries, SN) || ''
        };
        option.__excelHidden = hiddenFlags;
        option.__excelRef = {
            categories: baseRefs.catRef || '',
            series: excelSeries
        };
    } else if (layoutId === 'clusteredColumn') {
        const isHistogram = hasBinning;
        const categories = rangeStrToCellRefs(baseRefs.catRef, SN);
        option.xAxis = { type: 'category', data: categories };
        option.yAxis = { type: 'value' };
        option.series = activeSeries.map((ser, index) => {
            const dataId = String(ser?.['cx:dataId']?._$val ?? index);
            const refs = dataMap.get(dataId) || {};
            const values = rangeStrToCellRefs(refs.valRef, SN);
            return {
                name: getChartExSeriesNameRef(ser, SN) || '\u7cfb\u5217' + (index + 1),
                type: 'bar',
                data: values
            };
        });
        if (isHistogram) {
            const binning = baseSeries?.['cx:layoutPr']?.['cx:binning'];
            option.__chartType = 'histogram';
            option.__histogram = {
                intervalClosed: binning?._$intervalClosed || 'r'
            };
        }
        option.__excelHidden = hiddenFlags;
        option.__excelRef = {
            categories: baseRefs.catRef || '',
            series: excelSeries
        };
    }

    return chartOptionApplyTemplate(option, SN);
}

export const refOptionParse = (option, SN) => {
    // Resolve refs through a JSON reviver to avoid manual string replacement bugs.
    const refRegex = /^(?:'[^']+'|[^'!]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;
    const ops = JSON.parse(JSON.stringify(option), (key, value) => {
        if (typeof value === 'string' && refRegex.test(value)) return getDataByRef(value, SN);
        return value;
    });

    trimPivotChartRefData(ops);

    // Re-normalize percent-stacked series after ref values are expanded.
    if (ops.__percent && Array.isArray(ops.series)) {
        const dataSeries = ops.series.filter(s => Array.isArray(s.data) && s.data.length > 0);
        if (dataSeries.length > 0) {
            const len = dataSeries[0].data.length;
            const totals = new Array(len).fill(0);
            dataSeries.forEach(s => s.data.forEach((val, idx) => { totals[idx] += (Number(val) || 0); }));
            dataSeries.forEach(s => {
                s.data = s.data.map((val, idx) => {
                    const total = totals[idx];
                    return total ? Math.round(((Number(val) || 0) / total) * 10000) / 100 : 0;
                });
            });
        }
    }

    // Final series cleanup.
    ops.series.forEach(s => {
        if (s.type === 'scatter') { // Excel scatter values may be stored as text.
            s.data = s.data.map((d, i) => {
                const x = Number(d[0]);
                const y = Number(d[1]);
                const size = Array.isArray(d) && d.length > 2 ? d[2] : undefined;
                const xVal = Number.isNaN(x) ? i : x;
                const yVal = Number.isNaN(y) ? i : y;
                return size !== undefined ? [xVal, yVal, size] : [xVal, yVal];
            });
            if (s.__bubble) {
                s.large = false;
                const sizes = s.data.map(item => {
                    if (Array.isArray(item)) return Number(item[2] ?? item[1] ?? item[0]);
                    return Number(item);
                }).filter(num => Number.isFinite(num));
                const minRaw = Number.isFinite(s.__bubbleMin) ? s.__bubbleMin : Math.min(...sizes);
                const maxRaw = Number.isFinite(s.__bubbleMax) ? s.__bubbleMax : Math.max(...sizes);
                const minSize = 6;
                const maxSize = 30;
                if (!Number.isFinite(minRaw) || !Number.isFinite(maxRaw) || maxRaw === minRaw) {
                    const fixed = (minSize + maxSize) / 2;
                    s.symbolSize = () => fixed;
                } else {
                    s.symbolSize = (val) => {
                        const raw = Array.isArray(val) ? (val[2] ?? val[1] ?? val[0]) : val;
                        const num = Number.isFinite(Number(raw)) ? Number(raw) : minRaw;
                        return minSize + (maxSize - minSize) * ((num - minRaw) / (maxRaw - minRaw));
                    };
                }
            }
        }
        applySeriesLabelFormatter(s);
    });

    if (ops.__percent) {
        applyPercentStackedAxis(ops.xAxis);
        applyPercentStackedAxis(ops.yAxis);
    }
    applyAxisFormatters(ops.xAxis);
    applyAxisFormatters(ops.yAxis);
    applyCategoryAxisLabelVisibility(ops.xAxis);
    applyCategoryAxisLabelVisibility(ops.yAxis);

    // Use a shared radar min/max so every indicator keeps the same scale.
    if (ops.series?.some(s => s.type === 'radar') && ops.radar?.indicator) {
        const allValues = [].concat(...ops.series
            .filter(s => s.type === 'radar')
            .flatMap(s => s.data.map(d => d.value))
        ).filter(v => typeof v === 'number' && !isNaN(v));

        if (allValues.length > 0) {
            const explicitMin = Number.isFinite(Number(ops.radar.__excelAxisMin)) ? Number(ops.radar.__excelAxisMin) : null;
            const explicitMax = Number.isFinite(Number(ops.radar.__excelAxisMax)) ? Number(ops.radar.__excelAxisMax) : null;
            const explicitInterval = Number.isFinite(Number(ops.radar.__excelAxisInterval)) ? Number(ops.radar.__excelAxisInterval) : null;
            let globalMin = explicitMin ?? (Math.min(...allValues) < 0 ? Math.min(...allValues) : 0);
            let globalMax = explicitMax ?? Math.max(...allValues);
            let interval = explicitInterval;
            if (!Number.isFinite(interval) || interval <= 0) {
                interval = getNiceRadarInterval((globalMax - globalMin || globalMax || 1) / 5);
            }
            if (explicitMin == null && globalMin >= 0) {
                globalMin = 0;
            } else if (explicitMin == null) {
                globalMin = Math.floor(globalMin / interval) * interval;
            }
            if (explicitMax == null) {
                globalMax = Math.ceil(globalMax / interval) * interval;
            }
            const splitNumber = Math.max(1, Math.round((globalMax - globalMin) / interval));
            ops.radar.splitNumber = splitNumber;
            ops.radar.indicator.forEach(ind => {
                if (ind.axisLabel == null) {
                    ind.axisLabel = { show: false };
                }
                if (ind.min == null) ind.min = globalMin;
                if (!ind.max) ind.max = globalMax;
            });
        }

        // Rotate indicators so the first item starts at 12 o'clock in clockwise order.
        const len = ops.radar.indicator.length;
        const offset = Math.floor(len / 4); // Shift from 3 o'clock to 12 o'clock.

        // Reverse first so the final order becomes clockwise.
        ops.radar.indicator.reverse();

        // Then rotate right to move the start point to 12 o'clock.
        ops.radar.indicator = [
            ...ops.radar.indicator.slice(-offset),
            ...ops.radar.indicator.slice(0, -offset)
        ];

        ops.series.forEach(s => {
            if (s.type === 'radar') {
                s.data.forEach(d => {
                    if (Array.isArray(d.value)) {
                        d.value.reverse();
                        d.value = [
                            ...d.value.slice(-offset),
                            ...d.value.slice(0, -offset)
                        ];
                    }
                });
            }
        });
    }

    if (ops.__chartType === 'waterfall' && ops.__waterfall) {
        const categories = Array.isArray(ops.__waterfall.categories) ? ops.__waterfall.categories : [];
        const valuesRaw = Array.isArray(ops.__waterfall.values) ? ops.__waterfall.values : [];
        const values = valuesRaw.map(val => {
            const num = Number(val);
            return Number.isFinite(num) ? num : 0;
        });
        const seriesName = ops.__waterfall.seriesName || ops.title?.text || '\u7011\u5e03';
        let cumulative = 0;
        const assist = [];
        const actual = values.map((val) => {
            const start = cumulative;
            cumulative += val;
            assist.push(start);
            return val;
        });
        const increaseColor = '#4472C4';
        const decreaseColor = '#ED7D31';
        const totalColor = '#FFC000';
        const colorFn = (params) => {
            const val = actual[params.dataIndex] ?? 0;
            return val >= 0 ? increaseColor : decreaseColor;
        };
        ops.xAxis = { type: 'category', data: categories };
        ops.yAxis = { type: 'value' };
        ops.xAxis.axisTick = { alignWithLabel: true };
        if (ops.legend?.show !== false) {
            ops.legend = {
                ...(ops.legend || {}),
                data: ['\u589e\u52a0', '\u51cf\u5c11', '\u603b\u8ba1'],
                selectedMode: false,
                top: ops.legend?.top ?? 40,
                left: ops.legend?.left ?? 'center'
            };
        }
        const legendStub = categories.map(() => null);
        ops.series = [
            {
                name: '\u8f85\u52a9',
                type: 'bar',
                stack: 'total',
                barWidth: '40%',
                itemStyle: { color: 'transparent' },
                emphasis: { disabled: true },
                data: assist
            },
            {
                name: seriesName,
                type: 'bar',
                stack: 'total',
                barWidth: '40%',
                itemStyle: { color: colorFn },
                label: { show: true, position: 'top', color: '#595959' },
                data: actual
            },
            {
                name: '\u589e\u52a0',
                type: 'bar',
                stack: 'total',
                barWidth: '40%',
                data: legendStub,
                silent: true,
                itemStyle: { color: increaseColor }
            },
            {
                name: '\u51cf\u5c11',
                type: 'bar',
                stack: 'total',
                barWidth: '40%',
                data: legendStub,
                silent: true,
                itemStyle: { color: decreaseColor }
            },
            {
                name: '\u603b\u8ba1',
                type: 'bar',
                stack: 'total',
                barWidth: '40%',
                data: legendStub,
                silent: true,
                itemStyle: { color: totalColor }
            }
        ];
    }

    if (ops.__chartType === 'histogram' && Array.isArray(ops.series)) {
        const baseSeries = ops.series[0];
        const values = Array.isArray(baseSeries?.data) ? baseSeries.data : [];
        const nums = values.map(val => {
            const num = Number(val);
            return Number.isFinite(num) ? num : null;
        }).filter(val => val != null);
        if (nums.length > 0) {
            const intervalClosed = ops.__histogram?.intervalClosed || 'r';
            const mean = nums.reduce((sum, v) => sum + v, 0) / nums.length;
            const variance = nums.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / nums.length;
            const std = Math.sqrt(variance);
            const range = Math.max(...nums) - Math.min(...nums);
            let rawWidth = std > 0 ? (3.5 * std / Math.cbrt(nums.length)) : (range || 1);
            if (!Number.isFinite(rawWidth) || rawWidth <= 0) rawWidth = range || 1;
            const width = (() => {
                const abs = Math.abs(rawWidth);
                if (abs < 1) {
                    const power = Math.pow(10, Math.floor(Math.log10(abs)));
                    const fraction = abs / power;
                    const step = fraction <= 1 ? 1 : (fraction <= 2 ? 2 : (fraction <= 5 ? 5 : 10));
                    return step * power;
                }
                if (abs < 10) return Math.ceil(abs);
                if (abs < 100) return Math.ceil(abs / 5) * 5;
                if (abs < 1000) return Math.ceil(abs / 10) * 10;
                return Math.ceil(abs / 100) * 100;
            })();
            const minVal = Math.min(...nums);
            let binCount = Math.max(1, Math.ceil((Math.max(...nums) - minVal) / width));
            let edges = Array.from({ length: binCount + 1 }, (_, i) => minVal + i * width);
            if (edges[edges.length - 1] < Math.max(...nums)) {
                edges = edges.concat(edges[edges.length - 1] + width);
            }
            binCount = edges.length - 1;
            const counts = new Array(binCount).fill(0);
            nums.forEach((val) => {
                let idx;
                if (intervalClosed === 'r') {
                    idx = val <= edges[0] ? 0 : Math.ceil((val - edges[0]) / width) - 1;
                } else {
                    idx = Math.floor((val - edges[0]) / width);
                }
                if (idx < 0) idx = 0;
                if (idx >= binCount) idx = binCount - 1;
                counts[idx] += 1;
            });
            const formatNumber = (num) => {
                if (!Number.isFinite(num)) return '';
                const rounded = Math.round(num * 1000) / 1000;
                if (Number.isInteger(rounded)) return String(rounded);
                return String(rounded).replace(/\.?0+$/, '');
            };
            const labels = counts.map((_, i) => {
                const left = formatNumber(edges[i]);
                const right = formatNumber(edges[i + 1]);
                if (intervalClosed === 'l') {
                    const rightBracket = i === binCount - 1 ? ']' : ')';
                    return `[${left}, ${right}${rightBracket}`;
                }
                if (i === 0) return `[${left}, ${right}]`;
                return `(${left}, ${right}]`;
            });
            ops.legend = { ...(ops.legend || {}), show: false };
            ops.xAxis = { type: 'category', data: labels };
            ops.xAxis.axisLabel = { ...(ops.xAxis.axisLabel || {}), margin: 4 };
            ops.yAxis = { type: 'value' };
            ops.series = [{
                name: baseSeries?.name || ops.title?.text || '\u76f4\u65b9\u56fe',
                type: 'bar',
                data: counts,
                barGap: '0%',
                barCategoryGap: '0%'
            }];
            ops.grid = { ...(ops.grid || {}), bottom: 5 };
        }
    }

    return ops

};



// Disable animation for lower overhead and predictable output.
export const chartOptionApplyTemplate = (option, SN = null) => {

    const optionTem = {
        color: [
            "#4472C4", // Blue (Accent1)
            "#ED7D31", // Orange (Accent2)
            "#FFC000", // Gold (Accent4)
            "#70AD47", // Green (Accent6)
            "#5B9BD5", // Light Blue (Accent5)
            "#A5A5A5", // Gray (Accent3)
            
            "#264478", // Dark Blue
            "#9E480E", // Dark Orange
            "#997300", // Dark Gold
            "#375623",  // Dark Green
            "#1F4E79", // Navy Blue
            "#636363", // Dark Gray
        ],
        backgroundColor: '#ffffff',
        title: {
            text: SN?.t ? SN.t('chart.defaultTitle') : 'Chart Title',
            left: 'center',
            top: 10,
            textStyle: {
                fontSize: 18,
                fontWeight: 'normal'
            }
        },
        legend: {
            orient: 'horizontal',
            left: 'center',
            textStyle: {
                fontSize: 14
            }
        }
    }

    if (!option.legend?.bottom && !option.legend?.top) option.legend = { bottom: 10, ...option.legend } // Default legend to the bottom.
    normalizeLegendAnchors(option.legend);

    // Apply xAxis defaults.
    if (option.xAxis) {
        if (Array.isArray(option.xAxis)) {
            option.xAxis.forEach((a, i) => {
                option.xAxis[i].axisTick = { show: false, ...(a.axisTick || {}) }
            })
        } else {
            option.xAxis.axisTick = { show: false, ...(option.xAxis.axisTick || {}) }
        }
    }

    // Apply yAxis defaults.
    if (option.yAxis) {
        if (Array.isArray(option.yAxis)) {
            option.yAxis.forEach((a, i) => {
                option.yAxis[i].axisTick = { show: false, ...(a.axisTick || {}) }
            })
        } else {
            option.yAxis.axisTick = { show: false, ...(option.yAxis.axisTick || {}) }
        }
    }

    // Force non-interactive preview defaults.
    option.tooltip = { trigger: 'none', show: false }
    option.animation = false
    option.animationDuration = 0
    option.silent = true

    option = mergeTemplate(optionTem, option)
    normalizeLegendAnchors(option.legend);
    applyCategoryAxisLabelVisibility(option.xAxis);
    applyCategoryAxisLabelVisibility(option.yAxis);
    const axisLayout = applyAxisOffsets(option);

    const type = {}
    if (Array.isArray(option.series)) option.series.forEach((s, index) => {
        type[s.type] = true
        s.emphasis = {
            ...(s.emphasis || {}),
            scale: false
        };
        s.large = true;
        s.largeThreshold = 0;
        s.progressive = 500;
        s.progressiveThreshold = 1000;
        if (s.type == 'radar') { // Radar defaults.
            const base = option.series[index] || {};
            const hasLabel = hasRadarDataLabel(base);
            option.radar = {
                center: option.legend.top ? ['50%', '63%'] : ['50%', '54%'],
                radius: '50%',
                ...option.radar
            };
            option.series[index] = {
                ...base,
                ...(base.symbolSize == null ? { symbolSize: hasLabel ? 0.01 : 0 } : {})
            };
        } else if (s.type == 'pie') { // Pie / doughnut defaults.
            option.series[index] = {
                radius: '50%',
                label: { show: false },
                labelLine: { show: false },
                ...option.series[index]
            }
        } else if (s.type == 'scatter') { // Scatter defaults.
            const xAxes = Array.isArray(option.xAxis) ? option.xAxis : (option.xAxis ? [option.xAxis] : []);
            xAxes.forEach((axis, axisIndex) => {
                xAxes[axisIndex].splitLine = { show: true, ...(axis?.splitLine || {}) };
            });
            const base = {
                symbolSize: 7.5,
                ...option.series[index]
            };
            base.large = false;
            option.series[index] = base;
        } else if (s.type == 'sunburst') {
            const base = option.series[index] || {};
            option.series[index] = {
                ...base,
                radius: base.radius || ['35%', '75%'],
                center: base.center || ['50%', '55%'],
                sort: base.sort ?? null,
                nodeClick: base.nodeClick ?? false,
                label: { show: true, rotate: 'radial', color: '#ffffff', ...(base.label || {}) },
                itemStyle: { borderWidth: 2, borderColor: '#ffffff', ...(base.itemStyle || {}) }
            };
        } else if (s.type == 'treemap') {
            const base = option.series[index] || {};
            const legendData = Array.isArray(option.legend?.data) && option.legend.data.length > 0
                ? option.legend.data
                : (Array.isArray(base.data) ? base.data.map(item => item?.name).filter(Boolean) : []);
            option.series[index] = {
                ...base,
                leafDepth: base.leafDepth ?? 1,
                label: { show: true, position: 'insideTopLeft', color: '#ffffff', ...(base.label || {}) },
                upperLabel: { show: false, ...(base.upperLabel || {}) },
                breadcrumb: { show: false, ...(base.breadcrumb || {}) },
                itemStyle: { borderWidth: 2, borderColor: '#ffffff', ...(base.itemStyle || {}) }
            };
            option.legend = {
                ...(option.legend || {}),
                show: true,
                top: option.legend?.top ?? 40,
                left: option.legend?.left ?? 'center',
                data: legendData
            };
        } else if (s.type == 'funnel') {
            option.series[index] = {
                sort: 'none',
                gap: 6,
                minSize: '90%',
                maxSize: '90%',
                funnelAlign: 'center',
                colorBy: 'series',
                itemStyle: { color: '#4472C4', ...(option.series[index]?.itemStyle || {}) },
                label: {
                    show: true,
                    position: 'inside',
                    formatter: '{c}',
                    color: '#ffffff',
                    align: 'center',
                    verticalAlign: 'middle'
                },
                labelLine: { show: false },
                top: 40,
                bottom: 20,
                ...option.series[index]
            };
        } else if (s.type == 'bar') {
            const base = option.series[index] || {};
            const hasLabel = base.label?.show === true;
            const hasOutsideTopLabel = hasLabel && (
                base.label?.position === 'top' ||
                base.label?.position === 'outside' ||
                base.label?.__excelPosition === 'outEnd'
            );
            option.series[index] = {
                ...base,
                large: hasLabel ? false : base.large,
                ...(hasOutsideTopLabel ? { clip: false } : {})
            };
            if (option.__chartType === 'histogram') {
                option.series[index] = {
                    barGap: '0%',
                    barCategoryGap: '0%',
                    ...option.series[index]
                };
            }
        } else if (s.type == 'line') {
            const base = option.series[index] || {};
            const hasLabel = base.label?.show === true;
            const showSymbol = base.showSymbol === true || hasLabel;
            const symbolSize = showSymbol
                ? (base.showSymbol === true ? (base.symbolSize ?? 7.5) : (base.symbolSize ?? 0))
                : 0;
            option.series[index] = {
                ...base,
                large: hasLabel ? false : base.large,
                ...(hasLabel && base.clip == null ? { clip: false } : {}),
                showSymbol,
                ...(hasLabel ? { showAllSymbol: true } : {}),
                symbolSize
            };
        }
    });
    // Match legend icon size to the chart type.
    const legendIconSize = {
        line: { width: 25, height: 3 },
        bar: { width: 8, height: 8 },
        scatter: { width: 6, height: 6 },
        pie: { width: 6, height: 6 },
        radar: { width: 25, height: 3 },
    };
    const seriesTypes = Object.keys(type);
    if (seriesTypes.length === 1 && legendIconSize[seriesTypes[0]]) {
        option.legend.itemWidth = legendIconSize[seriesTypes[0]].width;
        option.legend.itemHeight = legendIconSize[seriesTypes[0]].height;
    } else {
        option.legend.itemWidth = legendIconSize.bar.width;
        option.legend.itemHeight = legendIconSize.bar.height;
    }

    const legendSideReserve = getLegendSideReserve(option);

    // Grid only affects axis-based charts.
    const gridTop = getDefaultGridTop(option, axisLayout.defaultTop);
    option.grid = {
        top: getGridOffsetValue(option.legend.top, 30, gridTop, 'top'),
        bottom: getGridOffsetValue(option.legend.bottom, 30, axisLayout.defaultBottom, 'bottom'),
        left: option.legend?.__excelSide && option.legend?.left === 'left'
            ? Math.max(axisLayout.defaultLeft, legendSideReserve)
            : axisLayout.defaultLeft,
        right: option.legend?.__excelSide && option.legend?.right != null
            ? Math.max(axisLayout.defaultRight, legendSideReserve)
            : axisLayout.defaultRight,
        containLabel: true,
        ...option.grid
    };

    return option
}

function getDefaultGridTop(option, fallback) {
    const hasTitle = option?.title?.show !== false && String(option?.title?.text ?? '').trim() !== '';
    const hasTopLegend = option?.legend?.show !== false && option?.legend?.top != null && option.legend.top !== 'middle';
    if (hasTitle || hasTopLegend) return fallback;
    return Math.max(28, fallback - 18);
}

function getLegendSideReserve(option) {
    const legend = option?.legend;
    if (!legend || legend.show === false || !legend.__excelSide) return 0;
    if (legend.right == null && legend.left !== 'left') return 0;

    const labels = getLegendLabels(option);
    const fontSize = Number(legend.textStyle?.fontSize) || 14;
    const itemWidth = Number(legend.itemWidth) || 8;
    const itemGap = Number(legend.itemGap) || 10;
    const maxTextWidth = Math.max(0, ...labels.map(label => measureLegendText(label, fontSize)));
    const reserve = itemWidth + itemGap + maxTextWidth + getLegendHorizontalPadding(legend.padding) + 14;
    return Math.min(240, Math.max(80, Math.ceil(reserve)));
}

function getLegendLabels(option) {
    const data = option?.legend?.data;
    if (Array.isArray(data) && data.length > 0) {
        return data.map(item => typeof item === 'object' ? item?.name : item).filter(item => item != null && item !== '').map(String);
    }

    const labels = [];
    (option?.series || []).forEach((series) => {
        const name = series?.name;
        if (name == null || name === '') return;
        const label = String(name);
        if (!labels.includes(label)) labels.push(label);
    });
    return labels;
}

function measureLegendText(text, fontSize) {
    let units = 0;
    for (const char of String(text)) {
        const code = char.codePointAt(0);
        if (code > 0x2e80) units += 0.85;
        else if (/[A-Z0-9]/.test(char)) units += 0.65;
        else if (/[a-z]/.test(char)) units += 0.55;
        else units += 0.45;
    }
    return Math.ceil(units * fontSize);
}

function getLegendHorizontalPadding(padding) {
    if (typeof padding === 'number' && Number.isFinite(padding)) return padding * 2;
    if (!Array.isArray(padding)) return 16;
    if (padding.length >= 4) return (Number(padding[1]) || 0) + (Number(padding[3]) || 0);
    if (padding.length >= 2) return (Number(padding[1]) || 0) * 2;
    return (Number(padding[0]) || 0) * 2;
}

function normalizeLegendAnchors(legend = {}) {
    if (legend.right != null) {
        delete legend.left;
    }
    if (legend.left === 'left') {
        delete legend.right;
    }
    if (legend.top === 'middle') {
        delete legend.bottom;
    }
    if (legend.bottom != null) {
        delete legend.top;
    }
}
// Merge data into a template object recursively.
function mergeTemplate(template, data) {
    for (const key in data) {
        const val = data[key];
        if (
            template.hasOwnProperty(key) &&
            isObject(template[key]) &&
            isObject(val)
        ) {
            // Merge nested objects recursively.
            mergeTemplate(template[key], val);
        } else {
            // Otherwise replace the target value.
            template[key] = val;
        }
    }
    return template;
}
function isObject(obj) {
    return obj !== null && typeof obj === 'object' && !Array.isArray(obj);
}

// Build legend options.
const getLegend = (legendObj, SN = null) => {
    if (!legendObj) return { show: false }
    // Base legend config.
    const legendOption = { show: true, orient: 'horizontal' };

    // Default legend position from c:legendPos.
    const posVal = legendObj['c:legendPos']?._$val;
    switch (posVal) {
        case 't':
            legendOption.top = 40;
            legendOption.left = 'center';
            break;
        case 'l':
            legendOption.left = 'left';
            legendOption.top = 'middle';
            legendOption.orient = 'vertical';
            legendOption.__excelSide = true;
            break;
        case 'r':
            legendOption.right = 'right';
            legendOption.top = 'middle';
            legendOption.orient = 'vertical';
            legendOption.__excelSide = true;
            break;
        default:
            legendOption.bottom = 10;
            legendOption.left = 'center';
    }

    // manualLayout overrides the default position.
    const manual = safeGet(legendObj, ['c:layout', 'c:manualLayout']);
    if (manual) {
        const xVal = manual['c:x']?._$val;
        const yVal = manual['c:y']?._$val;
        if (xVal != null) {
            legendOption.left = `${xVal * 100}%`;
            // Remove the opposite anchor to avoid conflicts.
            delete legendOption.right;
            delete legendOption.__excelSide;
        }
        if (yVal != null) {
            legendOption.top = `${yVal * 100}%`;
            // Remove the opposite anchor to avoid conflicts.
            delete legendOption.bottom;
            delete legendOption.__excelSide;
        }
    }

    // Font size.
    const defRPr = safeGet(legendObj, ['c:txPr', 'a:p', 'a:pPr', 'a:defRPr']);
    if (defRPr && defRPr._$sz) {
        legendOption.textStyle = legendOption.textStyle || {};
        legendOption.textStyle.fontSize = Number(defRPr._$sz) / 100 * 1.333;
    }
    const legendColor = getSolidFillColor(defRPr?.['a:solidFill'], SN);
    if (legendColor) {
        legendOption.textStyle = legendOption.textStyle || {};
        legendOption.textStyle.color = legendColor;
    }
    const legendStyle = getShapeStyleOption(legendObj?.['c:spPr'], SN);
    if (legendStyle.fillColor) {
        legendOption.backgroundColor = legendStyle.fillColor;
        legendOption.padding = legendOption.padding || [4, 8];
    }
    if (legendStyle.lineStyle?.show === true) {
        if (legendStyle.lineStyle.color) legendOption.borderColor = legendStyle.lineStyle.color;
        if (legendStyle.lineStyle.width != null) legendOption.borderWidth = legendStyle.lineStyle.width;
    }

    return legendOption
}

// Build title options.
const getTitle = (titleObj, SN = null) => {
    const obj = {
        show: false,
        text: '',
        textStyle: {}
    }
    // Title text.
    const titleRich = safeGet(titleObj, ['c:tx', 'c:rich', 'a:p', 'a:r']);
    if (titleRich) {
        if (Array.isArray(titleRich)) {
            obj.text = titleRich.map(p => p['a:t']).join('');
        } else if (titleRich['a:t']) {
            obj.text = titleRich['a:t']
        }
        obj.show = obj.text !== '';
    }
    // Title style.
    const style = titleRich?.["a:rPr"] || safeGet(titleObj, ['c:tx', 'c:rich', 'a:p', 'a:pPr', 'a:defRPr'])
    if (style) {
        obj.textStyle.fontSize = style?.["_$sz"] / 100 * 1.333 || 18
        obj.textStyle.fontWeight = style?.["_$b"] == '1' ? 'bold' : 'normal'
        style?.["_$i"] == '1' ? obj.textStyle.fontStyle = 'italic' : ''
        const textColor = getSolidFillColor(style?.['a:solidFill'], SN);
        if (textColor) obj.textStyle.color = textColor;
    }
    // Manual position.
    const layout = safeGet(titleObj, ['c:layout', 'c:manualLayout']);
    if (layout) {
        const xVal = layout['c:x']?._$val; // c:x and c:y are chart-area ratios in the 0..1 range.
        const yVal = layout['c:y']?._$val;
        if (xVal != null) { // Convert to percentages, e.g. 0.1 -> "10%".
            obj.left = `${xVal * 100}%`;
        }
        if (yVal != null) {
            obj.top = `${yVal * 100}%`;
        }
    }
    return obj
}

function normalizeHexColor(value) {
    if (typeof value !== 'string') return null;
    const hex = value.trim();
    if (!hex) return null;
    return hex.startsWith('#') ? hex.toUpperCase() : `#${hex.toUpperCase()}`;
}

function hexToRgba(color, alpha = 1) {
    const normalized = normalizeHexColor(color);
    if (!normalized) return undefined;
    const hex = normalized.slice(1);
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    const safeAlpha = Math.max(0, Math.min(1, alpha));
    if (safeAlpha >= 0.999) return normalized;
    return `rgba(${r}, ${g}, ${b}, ${Math.round(safeAlpha * 1000) / 1000})`;
}

function getColorAlpha(colorObj) {
    const alphaVal = Number(colorObj?.['a:alpha']?._$val);
    if (!Number.isFinite(alphaVal)) return 1;
    return Math.max(0, Math.min(1, alphaVal / 100000));
}

function getDrawingColor(colorObj, SN = null) {
    if (!colorObj || typeof colorObj !== 'object') return undefined;
    const color = getColor(colorObj, SN);
    if (!color) return undefined;
    return hexToRgba(color, getColorAlpha(colorObj));
}

function getSolidFillColor(fillObj, SN = null) {
    return getDrawingColor(fillObj?.['a:srgbClr'] || fillObj?.['a:schemeClr'] || fillObj?.['a:sysClr'], SN);
}

function isTransparentColor(fillObj) {
    const colorObj = fillObj?.['a:srgbClr'] || fillObj?.['a:schemeClr'] || fillObj?.['a:sysClr'];
    return !!colorObj && getColorAlpha(colorObj) <= 0;
}

function getSeriesLineWidth(ser) {
    const widthEmu = Number(safeGet(ser, ['c:spPr', 'a:ln', '_$w']));
    if (!Number.isFinite(widthEmu) || widthEmu <= 0) return undefined;
    return Math.round((widthEmu / 9525) * 100) / 100;
}

function getSeriesStrokeColor(ser, SN = null) {
    return getSolidFillColor(safeGet(ser, ['c:spPr', 'a:ln', 'a:solidFill']), SN);
}

function getSeriesFillColor(ser, SN = null) {
    return getSolidFillColor(safeGet(ser, ['c:spPr', 'a:solidFill']), SN) || getSeriesStrokeColor(ser, SN);
}

function toFiniteNumber(value) {
    const num = Number(value);
    return Number.isFinite(num) ? num : undefined;
}

function getLineDashType(ser) {
    const dash = safeGet(ser, ['c:spPr', 'a:ln', 'a:prstDash', '_$val']);
    if (typeof dash !== 'string' || !dash) return undefined;
    if (/dot/i.test(dash)) return 'dotted';
    if (/dash/i.test(dash)) return 'dashed';
    return undefined;
}

function getMarkerSymbol(markerObj) {
    const symbol = markerObj?.['c:symbol']?._$val;
    switch (symbol) {
        case 'circle':
        case 'diamond':
        case 'triangle':
        case 'star':
            return symbol;
        case 'square':
        case 'dash':
            return 'rect';
        case 'dot':
            return 'circle';
        default:
            return undefined;
    }
}

function getPointItemStyle(pointObj, SN = null) {
    const fillColor = getSolidFillColor(safeGet(pointObj, ['c:spPr', 'a:solidFill']), SN);
    const borderColor = getSolidFillColor(safeGet(pointObj, ['c:spPr', 'a:ln', 'a:solidFill']), SN);
    const borderWidth = getSeriesLineWidth(pointObj);
    if (!fillColor && !borderColor && borderWidth == null) return null;
    return {
        ...(fillColor ? { color: fillColor } : {}),
        ...(borderColor ? { borderColor } : {}),
        ...(borderWidth != null ? { borderWidth } : {})
    };
}

function getDrawingLineStyleOption(ln, SN = null) {
    if (!ln || typeof ln !== 'object') return null;
    if (ln['a:noFill'] !== undefined) return { show: false };
    const lineStyle = { show: true };
    const color = getSolidFillColor(ln?.['a:solidFill'], SN);
    const width = Number(ln?._$w);
    const dash = ln?.['a:prstDash']?._$val;
    if (color) lineStyle.color = color;
    if (Number.isFinite(width) && width > 0) {
        lineStyle.width = Math.round((width / 9525) * 100) / 100;
    }
    if (/dot/i.test(dash || '')) lineStyle.type = 'dotted';
    else if (/dash/i.test(dash || '')) lineStyle.type = 'dashed';
    return lineStyle;
}

function getShapeStyleOption(spPr, SN = null) {
    if (!spPr || typeof spPr !== 'object') return {};
    const option = {};
    if (spPr['a:noFill'] === undefined) {
        const fillColor = getSolidFillColor(spPr?.['a:solidFill'], SN);
        if (fillColor) option.fillColor = fillColor;
    }
    const lineStyle = getDrawingLineStyleOption(spPr?.['a:ln'], SN);
    if (lineStyle) option.lineStyle = lineStyle;
    return option;
}

function mergeLabelNodes(baseNode, overrideNode) {
    return { ...(baseNode || {}), ...(overrideNode || {}) };
}

function getTextStyleFromTxPr(txPr, SN = null, fallbackSize = 12) {
    const defRPr = safeGet(txPr, ['a:p', 'a:pPr', 'a:defRPr']) || safeGet(txPr, ['a:p', 'a:endParaRPr']);
    if (!defRPr) return {};
    const textStyle = {};
    const fontSize = Number(defRPr?._$sz);
    if (Number.isFinite(fontSize) && fontSize > 0) {
        textStyle.fontSize = Math.round((fontSize / 75) * 100) / 100;
    } else {
        textStyle.fontSize = fallbackSize;
    }
    const color = getSolidFillColor(defRPr?.['a:solidFill'], SN);
    if (color) textStyle.color = color;
    return textStyle;
}

function getAxisLineOption(axisObj, SN = null) {
    const ln = safeGet(axisObj, ['c:spPr', 'a:ln']);
    if (!ln || typeof ln !== 'object') return {};
    if (ln['a:noFill'] !== undefined) return { axisLine: { show: false } };
    const lineStyle = {};
    const color = getSolidFillColor(ln?.['a:solidFill'], SN);
    const width = Number(ln?._$w);
    const dash = ln?.['a:prstDash']?._$val;
    if (color) lineStyle.color = color;
    if (Number.isFinite(width) && width > 0) {
        lineStyle.width = Math.round((width / 9525) * 100) / 100;
    }
    if (/dot/i.test(dash || '')) lineStyle.type = 'dotted';
    else if (/dash/i.test(dash || '')) lineStyle.type = 'dashed';
    return { axisLine: { show: true, ...(Object.keys(lineStyle).length > 0 ? { lineStyle } : {}) } };
}

function getSplitLineStyleOption(axisObj, SN = null) {
    const gridLn = safeGet(axisObj, ['c:majorGridlines', 'c:spPr', 'a:ln']);
    const splitLine = { show: !!axisObj?.['c:majorGridlines'] };
    if (!gridLn || typeof gridLn !== 'object') return splitLine;
    const lineStyle = {};
    const color = getSolidFillColor(gridLn?.['a:solidFill'], SN);
    const width = Number(gridLn?._$w);
    const dash = gridLn?.['a:prstDash']?._$val;
    if (color) lineStyle.color = color;
    if (Number.isFinite(width) && width > 0) {
        lineStyle.width = Math.round((width / 9525) * 100) / 100;
    }
    if (/dot/i.test(dash || '')) lineStyle.type = 'dotted';
    else if (/dash/i.test(dash || '')) lineStyle.type = 'dashed';
    if (Object.keys(lineStyle).length > 0) splitLine.lineStyle = lineStyle;
    return splitLine;
}

function getLabelFlag(dLbls, key) {
    return dLbls?.[`c:${key}`]?._$val === '1';
}

function mapExcelLabelPosition(pos, seriesType) {
    if (!pos) return undefined;
    if (seriesType === 'pie') {
        if (pos === 'outEnd' || pos === 'bestFit') return 'outside';
        if (pos === 'inEnd') return 'inside';
        if (pos === 'ctr') return 'inside';
        if (pos === 'inBase') return 'inside';
    }
    switch (pos) {
        case 't': return 'top';
        case 'b': return 'bottom';
        case 'l': return 'left';
        case 'r': return 'right';
        case 'ctr': return 'inside';
        case 'inEnd': return 'insideTop';
        case 'inBase': return 'insideBottom';
        case 'outEnd': return 'top';
        case 'bestFit': return seriesType === 'pie' ? 'outside' : 'top';
        default: return undefined;
    }
}

function buildExcelLabelTemplate(parts, seriesType) {
    const lines = [];
    if (parts.showSeriesName) lines.push('{a}');
    if (parts.showCategoryName) lines.push('{b}');
    if (parts.showValue) lines.push('{c}');
    if (parts.showPercent) lines.push(seriesType === 'pie' ? '{d}%' : '{c}%');
    return lines.length > 0 ? lines.join('\n') : '';
}

function buildSeriesLabelOption(dLbls, seriesType, SN = null) {
    if (!dLbls) return {};
    const parts = {
        showSeriesName: getLabelFlag(dLbls, 'showSerName'),
        showCategoryName: getLabelFlag(dLbls, 'showCatName'),
        showValue: getLabelFlag(dLbls, 'showVal'),
        showPercent: getLabelFlag(dLbls, 'showPercent'),
        showBubbleSize: getLabelFlag(dLbls, 'showBubbleSize')
    };
    const show = Object.values(parts).some(Boolean);
    const position = mapExcelLabelPosition(dLbls?.['c:dLblPos']?._$val, seriesType);
    const formatter = buildExcelLabelTemplate(parts, seriesType);
    const hasLeaderLine = dLbls?.['c:showLeaderLines']?._$val === '1';
    if (!show && !position && !hasLeaderLine) return {};

    const label = {
        show,
        ...(position ? { position } : {})
    };
    Object.assign(label, getTextStyleFromTxPr(dLbls?.['c:txPr'], SN, 12));
    if (typeof dLbls?.['c:dLblPos']?._$val === 'string' && dLbls['c:dLblPos']._$val) {
        label.__excelPosition = dLbls['c:dLblPos']._$val;
    }
    if (formatter) label.formatter = formatter;
    if (show) label.__excelParts = parts;

    const result = { label };
    if (seriesType === 'pie') {
        result.labelLine = { show: hasLeaderLine };
    }
    return result;
}

function getAxisScaleOption(axisObj) {
    const scaling = axisObj?.['c:scaling'];
    const option = {};
    const min = toFiniteNumber(scaling?.['c:min']?._$val);
    const max = toFiniteNumber(scaling?.['c:max']?._$val);
    const logBase = toFiniteNumber(scaling?.['c:logBase']?._$val);
    const majorUnit = toFiniteNumber(axisObj?.['c:majorUnit']?._$val);
    const minorUnit = toFiniteNumber(axisObj?.['c:minorUnit']?._$val);
    const crossesAt = toFiniteNumber(axisObj?.['c:crossesAt']?._$val);
    const orientation = scaling?.['c:orientation']?._$val;
    const crosses = axisObj?.['c:crosses']?._$val;
    const crossBetween = axisObj?.['c:crossBetween']?._$val;
    if (min != null) option.min = min;
    if (max != null) option.max = max;
    if (logBase != null) {
        option.type = 'log';
        if (logBase !== 10) option.logBase = logBase;
    }
    if (majorUnit != null) option.interval = majorUnit;
    if (minorUnit != null) {
        option.minorTick = { show: true };
        option.__excelMinorUnit = minorUnit;
    }
    option.splitLine = { show: !!axisObj?.['c:majorGridlines'] };
    if (orientation === 'maxMin') option.inverse = true;
    if (crossesAt != null) option.__excelCrossesAt = crossesAt;
    if (typeof crosses === 'string' && crosses) {
        option.__excelCrosses = crosses;
        if (crosses === 'min') option.boundaryGap = false;
    }
    if (typeof crossBetween === 'string' && crossBetween) option.__excelCrossBetween = crossBetween;
    return option;
}

function hasPercentFormat(formatCode) {
    return typeof formatCode === 'string' && formatCode.includes('%');
}

function formatChartAxisValue(value, formatCode, options = {}) {
    if (value == null || value === '') return '';
    if (typeof formatCode !== 'string' || !formatCode || formatCode === 'General') return String(value);
    const numericValue = typeof value === 'number' ? value : Number(value);
    if (!Number.isFinite(numericValue)) return String(value);
    const formatValue = options.percentBase && hasPercentFormat(formatCode)
        ? numericValue / options.percentBase
        : numericValue;
    const formatted = numFmtFun(formatValue, formatCode);
    return typeof formatted === 'string' ? formatted : String(formatted ?? '');
}

function normalizeAxisOptionList(axisOption) {
    if (Array.isArray(axisOption)) return axisOption;
    return axisOption ? [axisOption] : [];
}

function applyCategoryAxisLabelVisibility(axisOption) {
    normalizeAxisOptionList(axisOption).forEach((axis) => {
        if (!axis || axis.type !== 'category') return;
        axis.axisLabel = axis.axisLabel || {};
        if (axis.axisLabel.show === false) return;

        axis.axisLabel.hideOverlap = axis.axisLabel.hideOverlap !== false;

        const denseLabels = shouldStaggerCategoryAxisLabels(axis);
        if (axis.axisLabel.interval == null) {
            axis.axisLabel.interval = denseLabels ? 'auto' : 0;
        }
        if (axis.axisLabel.rotate == null && shouldRotateCategoryAxisLabels(axis)) {
            axis.axisLabel.rotate = denseLabels ? 45 : 30;
        }
    });
}

function shouldStaggerCategoryAxisLabels(axis) {
    const data = Array.isArray(axis.data) ? axis.data : [];
    if (data.length <= 1) return false;

    const lengths = data.map(value => getCategoryAxisLabelLength(value));
    const maxLength = Math.max(0, ...lengths);
    const avgLength = lengths.reduce((sum, length) => sum + length, 0) / lengths.length;

    return data.length >= 12 || (data.length >= 8 && avgLength >= 3.5) || (data.length >= 5 && maxLength >= 7);
}

function shouldRotateCategoryAxisLabels(axis) {
    const data = Array.isArray(axis.data) ? axis.data : [];
    if (data.length === 0) return false;
    if (data.length >= 8) return true;

    const maxLength = Math.max(0, ...data.map(value => getCategoryAxisLabelLength(value)));
    return data.length >= 5 && maxLength >= 4;
}

function getCategoryAxisLabelLength(value) {
    if (value == null) return 0;
    const text = typeof value === 'object' ? (value.name ?? value.value ?? '') : value;
    return Array.from(String(text)).reduce((total, char) => {
        const code = char.codePointAt(0);
        return total + (code > 0x2e80 ? 2 : 1);
    }, 0);
}

function isBlankChartValue(value) {
    if (Array.isArray(value)) return value.every(isBlankChartValue);
    if (value && typeof value === 'object') return isBlankChartValue(value.value);
    return value === null || value === undefined || value === '' || (typeof value === 'number' && Number.isNaN(value));
}

function trimPivotChartRefData(option) {
    if (!option?.__pivotChart || !Array.isArray(option.series)) return;

    option.series.forEach((series) => {
        if (series?.type !== 'pie' || !Array.isArray(series.data)) return;
        let end = series.data.length;
        while (end > 0) {
            const item = series.data[end - 1];
            if (!isBlankChartValue(item?.name) || !isBlankChartValue(item?.value)) break;
            end--;
        }
        if (end < series.data.length) series.data = series.data.slice(0, end);
    });

    const trimAxis = (axisOption, axisType) => {
        normalizeAxisOptionList(axisOption).forEach((axis, axisIndex) => {
            if (!axis || axis.type !== 'category' || !Array.isArray(axis.data)) return;
            const linkedSeries = option.series.filter(series => {
                if (!Array.isArray(series?.data)) return false;
                const seriesAxisIndex = axisType === 'x'
                    ? Number(series.xAxisIndex ?? 0)
                    : Number(series.yAxisIndex ?? 0);
                return seriesAxisIndex === axisIndex;
            });
            if (linkedSeries.length === 0) return;

            let end = axis.data.length;
            while (end > 0) {
                const index = end - 1;
                const hasSeriesValue = linkedSeries.some(series => !isBlankChartValue(series.data[index]));
                if (!isBlankChartValue(axis.data[index]) || hasSeriesValue) break;
                end--;
            }
            if (end === axis.data.length) return;

            axis.data = axis.data.slice(0, end);
            linkedSeries.forEach(series => {
                series.data = series.data.slice(0, end);
            });
        });
    };

    trimAxis(option.xAxis, 'x');
    trimAxis(option.yAxis, 'y');
}

function applyPercentStackedAxis(axisOption) {
    normalizeAxisOptionList(axisOption).forEach((axis) => {
        if (!axis || axis.type !== 'value') return;

        const min = Number(axis.min);
        const max = Number(axis.max);
        const interval = Number(axis.interval);
        if (Number.isFinite(min) && Math.abs(min) <= 1) axis.min = min * 100;
        else if (axis.min == null) axis.min = 0;
        if (Number.isFinite(max) && Math.abs(max) <= 1) axis.max = max * 100;
        else if (axis.max == null) axis.max = 100;
        if (Number.isFinite(interval) && interval > 0 && interval <= 1) axis.interval = interval * 100;
        else if (axis.interval == null && axis.min === 0 && axis.max === 100) axis.interval = 20;

        axis.__excelPercentBase = 100;
        if (!axis.__excelFormatCode || axis.__excelFormatCode === 'General') {
            axis.__excelFormatCode = '0%';
        }
    });
}

function applyAxisFormatters(axisOption) {
    normalizeAxisOptionList(axisOption).forEach((axis) => {
        const formatCode = axis?.__excelFormatCode;
        if (typeof formatCode !== 'string' || !formatCode || formatCode === 'General') return;
        axis.axisLabel = axis.axisLabel || {};
        axis.axisLabel.formatter = (value) => formatChartAxisValue(value, formatCode, {
            percentBase: axis.__excelPercentBase
        });
    });
}

function applyAxisOffsets(option) {
    const resolveDefaultPosition = (axisType, index) => {
        if (axisType === 'x') return index % 2 === 0 ? 'bottom' : 'top';
        return index % 2 === 0 ? 'left' : 'right';
    };

    const assignOffsets = (axes, axisType, fallbackPosition) => {
        const offsetCounter = new Map();
        axes.forEach((axis, axisIndex) => {
            if (!axis || typeof axis !== 'object') return;
            const position = axis.position || resolveDefaultPosition(axisType, axisIndex) || fallbackPosition;
            axis.position = position;
            if (axis.show === false) return;
            const index = offsetCounter.get(position) || 0;
            if (axis.offset == null && index > 0) {
                axis.offset = index * 40;
            }
            offsetCounter.set(position, index + 1);
        });
        return offsetCounter;
    };

    const xAxes = normalizeAxisOptionList(option.xAxis);
    const yAxes = normalizeAxisOptionList(option.yAxis);
    assignOffsets(xAxes, 'x', 'bottom');
    assignOffsets(yAxes, 'y', 'left');

    const countByPosition = (axes, position) => axes.filter(axis => axis?.show !== false && (axis?.position || '') === position).length;
    const defaultLeft = 20 + Math.max(0, countByPosition(yAxes, 'left') - 1) * 40;
    const defaultRight = 20 + Math.max(0, countByPosition(yAxes, 'right') - 1) * 40;
    const defaultTop = 50 + Math.max(0, countByPosition(xAxes, 'top')) * 25;
    const defaultBottom = 10 + Math.max(0, countByPosition(xAxes, 'bottom') - 1) * 25;

    return { defaultLeft, defaultRight, defaultTop, defaultBottom };
}

function getGridOffsetValue(baseValue, delta, fallback, axis = 'top') {
    if (typeof baseValue === 'number' && Number.isFinite(baseValue)) return baseValue + delta;
    if (baseValue === 'middle') return fallback;
    if (axis === 'top' && baseValue === 'top') return fallback + delta;
    if (axis === 'bottom' && baseValue === 'bottom') return fallback + delta;
    if (typeof baseValue === 'string' && baseValue.trim()) return baseValue;
    return fallback;
}

function applyPivotChartPresentation(option, chartSpace, plotArea, SN = null) {
    const dataFieldName = getPivotChartDataFieldName(plotArea);
    if (dataFieldName) {
        const visibleSeries = option.series?.filter(series => series?.name) || [];
        const pieLegendData = getPivotPieLegendData(option);
        if (pieLegendData.length > 0) {
            option.series?.forEach(series => {
                if (series?.type === 'pie' && (!series.name || isRefWithSheet(series.name))) {
                    series.name = dataFieldName;
                }
            });
            option.legend = {
                ...(option.legend || {}),
                show: option.legend?.show !== false,
                data: pieLegendData
            };
            return;
        }

        if (visibleSeries.length <= 1) {
            option.series?.forEach(series => {
                if (!series.name || isRefWithSheet(series.name)) series.name = dataFieldName;
            });
            option.legend = {
                ...(option.legend || {}),
                show: option.legend?.show !== false,
                data: option.legend?.data || [dataFieldName]
            };
        } else {
            option.legend = {
                ...(option.legend || {}),
                show: option.legend?.show !== false
            };
        }
    }
}

function getPivotPieLegendData(option) {
    const pieSeries = option.series?.find(series => series?.type === 'pie' && Array.isArray(series.data));
    if (!pieSeries) return [];
    return pieSeries.data
        .map(item => item && typeof item === 'object' ? item.name : '')
        .filter(item => !isBlankChartValue(item));
}

function getPivotChartDataFieldName(plotArea) {
    const chartGroups = [
        ...sureArray(plotArea?.['c:barChart']),
        ...sureArray(plotArea?.['c:bar3DChart']),
        ...sureArray(plotArea?.['c:lineChart']),
        ...sureArray(plotArea?.['c:line3DChart']),
        ...sureArray(plotArea?.['c:areaChart']),
        ...sureArray(plotArea?.['c:area3DChart']),
        ...sureArray(plotArea?.['c:pieChart']),
        ...sureArray(plotArea?.['c:pie3DChart']),
        ...sureArray(plotArea?.['c:doughnutChart'])
    ];
    const firstSeries = chartGroups.flatMap(group => sureArray(group?.['c:ser'])).find(Boolean);
    const cached = getCachedPoints(safeGet(firstSeries, ['c:tx', 'c:strRef']))[0];
    return cached || '';
}

function hasRadarDataLabel(series) {
    if (series?.label?.show === true) return true;
    return Array.isArray(series?.data) && series.data.some(item => item?.label?.show === true);
}

function getNiceRadarInterval(rawStep) {
    if (!Number.isFinite(rawStep) || rawStep <= 0) return 1;
    const magnitude = 10 ** Math.floor(Math.log10(rawStep));
    const normalized = rawStep / magnitude;
    if (normalized <= 1) return magnitude;
    if (normalized <= 2) return 2 * magnitude;
    if (normalized <= 2.5) return 2.5 * magnitude;
    if (normalized <= 5) return 5 * magnitude;
    return 10 * magnitude;
}

function applySeriesLabelFormatter(series) {
    const parts = series?.label?.__excelParts;
    if (!parts) return;
    if (!parts.showBubbleSize) return;
    series.label.formatter = (params) => {
        const lines = [];
        if (parts.showSeriesName) lines.push(params.seriesName ?? series.name ?? '');
        if (parts.showCategoryName) lines.push(params.name ?? '');
        if (parts.showValue) {
            const val = Array.isArray(params.value) ? (params.value[1] ?? params.value[0] ?? '') : params.value;
            lines.push(val);
        }
        if (parts.showPercent && typeof params.percent === 'number') {
            lines.push(`${params.percent}%`);
        }
        if (parts.showBubbleSize) {
            const bubbleSize = Array.isArray(params.value) ? (params.value[2] ?? params.value[1] ?? params.value[0] ?? '') : params.value;
            lines.push(bubbleSize);
        }
        return lines.filter(item => item !== '' && item != null).join('\n');
    };
}

function getAxisId(axisObj) {
    const id = axisObj?.['c:axId']?._$val;
    return id == null ? '' : String(id);
}

function cloneChartXmlNode(node) {
    if (!node || typeof node !== 'object') return node;
    return JSON.parse(JSON.stringify(node));
}

function getAxisPosition(axisObj, fallback) {
    const pos = axisObj?.['c:axPos']?._$val;
    if (pos === 't') return 'top';
    if (pos === 'b') return 'bottom';
    if (pos === 'l') return 'left';
    if (pos === 'r') return 'right';
    return fallback;
}

function buildAxisTitleOption(axisObj, SN = null) {
    const title = getTitle(axisObj?.['c:title'], SN);
    if (!title?.text) return {};
    const result = { name: title.text };
    if (title.textStyle && Object.keys(title.textStyle).length > 0) {
        result.nameTextStyle = title.textStyle;
    }
    return result;
}

function getAxisFontSize(axisObj, fallback = 11) {
    const size = Number(safeGet(axisObj, ['c:txPr', 'a:p', 'a:pPr', 'a:defRPr', '_$sz']));
    if (!Number.isFinite(size) || size <= 0) return fallback;
    return Math.round((size / 75) * 100) / 100;
}

function getAxisLabelColor(axisObj, SN = null) {
    return getSolidFillColor(safeGet(axisObj, ['c:txPr', 'a:p', 'a:pPr', 'a:defRPr', 'a:solidFill']), SN);
}

function isAxisHidden(axisObj) {
    return axisObj?.['c:delete']?._$val === '1' || axisObj?.['c:tickLblPos']?._$val === 'none';
}

function buildCategoryAxisOption(axisObj, SN = null, fallbackPosition = 'bottom') {
    const axisLabel = { fontSize: getAxisFontSize(axisObj, 11) };
    const color = getAxisLabelColor(axisObj, SN);
    const hidden = isAxisHidden(axisObj);
    const transparentLabel = isTransparentColor(safeGet(axisObj, ['c:txPr', 'a:p', 'a:pPr', 'a:defRPr', 'a:solidFill']));
    const sourceLinked = axisObj?.['c:numFmt']?._$sourceLinked;
    if (color) axisLabel.color = color;
    if (transparentLabel) {
        axisLabel.show = false;
    }
    const option = {
        type: 'category',
        axisLabel,
        ...(hidden ? { show: false } : {}),
        ...(transparentLabel ? { __excelTransparentAxisLabel: true } : {}),
        position: getAxisPosition(axisObj, fallbackPosition),
        ...buildAxisTitleOption(axisObj, SN),
        ...getAxisScaleOption(axisObj),
        ...getAxisLineOption(axisObj, SN),
        splitLine: getSplitLineStyleOption(axisObj, SN)
    };
    const formatCode = axisObj?.['c:numFmt']?._$formatCode;
    if (typeof formatCode === 'string' && formatCode && formatCode !== 'General') {
        option.__excelFormatCode = formatCode;
    }
    if (sourceLinked === '0' || sourceLinked === '1') {
        option.__excelSourceLinked = sourceLinked;
    }
    if (transparentLabel) {
        const rawTxPr = cloneChartXmlNode(axisObj?.['c:txPr']);
        const rawSpPr = cloneChartXmlNode(axisObj?.['c:spPr']);
        if (rawTxPr) option.__excelRawAxisTxPr = rawTxPr;
        if (rawSpPr) option.__excelRawAxisSpPr = rawSpPr;
    }
    return option;
}

function buildValueAxisOption(axisObj, SN = null, fallbackPosition = 'left') {
    const axisLabel = { fontSize: getAxisFontSize(axisObj, 11) };
    const color = getAxisLabelColor(axisObj, SN);
    const hidden = isAxisHidden(axisObj);
    const transparentLabel = isTransparentColor(safeGet(axisObj, ['c:txPr', 'a:p', 'a:pPr', 'a:defRPr', 'a:solidFill']));
    const sourceLinked = axisObj?.['c:numFmt']?._$sourceLinked;
    if (color) axisLabel.color = color;
    if (transparentLabel) {
        axisLabel.show = false;
    }
    const formatCode = axisObj?.['c:numFmt']?._$formatCode;
    const option = {
        type: 'value',
        axisLabel,
        ...(hidden ? { show: false } : {}),
        ...(transparentLabel ? { __excelTransparentAxisLabel: true } : {}),
        position: getAxisPosition(axisObj, fallbackPosition),
        ...buildAxisTitleOption(axisObj, SN),
        ...getAxisScaleOption(axisObj),
        ...getAxisLineOption(axisObj, SN),
        splitLine: getSplitLineStyleOption(axisObj, SN)
    };
    if (typeof formatCode === 'string' && formatCode) {
        option.__excelFormatCode = formatCode;
    }
    if (sourceLinked === '0' || sourceLinked === '1') {
        option.__excelSourceLinked = sourceLinked;
    }
    if (transparentLabel) {
        const rawTxPr = cloneChartXmlNode(axisObj?.['c:txPr']);
        const rawSpPr = cloneChartXmlNode(axisObj?.['c:spPr']);
        if (rawTxPr) option.__excelRawAxisTxPr = rawTxPr;
        if (rawSpPr) option.__excelRawAxisSpPr = rawSpPr;
    }
    return option;
}

function buildRadarCoordinateOption(indicators, categoryAxisObj, valueAxisObj, SN = null) {
    const categoryTextStyle = getTextStyleFromTxPr(categoryAxisObj?.['c:txPr'], SN, 12);
    const valueTextStyle = getTextStyleFromTxPr(valueAxisObj?.['c:txPr'], SN, 12);
    const valueLabelHidden = isAxisHidden(valueAxisObj) || isTransparentColor(safeGet(valueAxisObj, ['c:txPr', 'a:p', 'a:pPr', 'a:defRPr', 'a:solidFill']));
    const valueScale = getAxisScaleOption(valueAxisObj);
    const valueFormatCode = valueAxisObj?.['c:numFmt']?._$formatCode;
    const firstAxisLabel = valueLabelHidden ? null : {
        show: true,
        ...valueTextStyle,
        ...(typeof valueFormatCode === 'string' && valueFormatCode && valueFormatCode !== 'General'
            ? { formatter: (value) => formatChartAxisValue(value, valueFormatCode) }
            : {})
    };
    const axisLine = getAxisLineOption(categoryAxisObj, SN).axisLine;
    const splitLine = getSplitLineStyleOption(valueAxisObj, SN);
    const radarOption = {
        indicator: indicators.map((name, index) => ({
            name,
            ...(Number.isFinite(valueScale.min) ? { min: valueScale.min } : {}),
            ...(Number.isFinite(valueScale.max) ? { max: valueScale.max } : {}),
            ...(index === 0 && firstAxisLabel ? { axisLabel: firstAxisLabel } : {})
        })),
        axisName: {
            show: !isAxisHidden(categoryAxisObj),
            ...categoryTextStyle
        },
        axisLabel: { show: false },
        splitArea: { show: false },
        ...(axisLine ? { axisLine } : {}),
        ...(splitLine ? { splitLine } : {})
    };
    if (Number.isFinite(valueScale.min)) radarOption.__excelAxisMin = valueScale.min;
    if (Number.isFinite(valueScale.max)) radarOption.__excelAxisMax = valueScale.max;
    if (Number.isFinite(valueScale.interval)) radarOption.__excelAxisInterval = valueScale.interval;
    return radarOption;
}

function attachAxisIndices(seriesObj, xAxisIndex, yAxisIndex) {
    if (xAxisIndex > 0) seriesObj.xAxisIndex = xAxisIndex;
    if (yAxisIndex > 0) seriesObj.yAxisIndex = yAxisIndex;
    return seriesObj;
}

// Normalize a value to an array.
const sureArray = (obj) => {
    if (!obj) return [];
    return Array.isArray(obj) ? obj : [obj];
};

function resolveDefinedNameRef(ref, SN) {
    if (typeof ref !== 'string') return '';
    const trimmed = ref.trim();
    const raw = trimmed.startsWith('=') ? trimmed.slice(1) : trimmed;
    if (raw.includes('!')) return raw;
    const defined = SN?._definedNames?.[raw];
    return typeof defined === 'string' ? defined : raw;
}

const getDataByRef = (ref, SN) => {
    if (typeof ref !== 'string') return '';
    const normalized = resolveDefinedNameRef(ref, SN);
    if (!normalized.includes('!')) return '';
    const refParts = normalized.split('!');
    const sheetName = refParts[0].replaceAll('\'', '');
    const sheetRef = refParts[1].replaceAll('$', '');
    const data = [];
    const sheet = SN.getSheet(sheetName);
    if (sheetRef.includes(':')) {
        sheet.eachCells(sheetRef, (r, c) => data.push(sheet.getCell(r, c).calcVal));
    } else {
        return sheet.getCell(sheetRef).calcVal
    }
    return data
};

// Safe nested property lookup.
const safeGet = (o, path) => {
    return path.reduce((acc, key) => (acc && acc[key] != null) ? acc[key] : undefined, o);
};

// Extract the series name.
const extractSeriesName = (ser) => {
    const txRef = safeGet(ser, ['c:tx', 'c:strRef', 'c:f']);
    // Guard against empty refs.
    if (!txRef || typeof txRef !== 'string') return '';
    const nameData = getDataByRef(txRef);
    if (Array.isArray(nameData)) return nameData[0] ?? '';
    return nameData ?? '';
};

function getSeriesName(ser, index, useCache = false) {
    const cached = useCache ? getCachedPoints(safeGet(ser, ['c:tx', 'c:strRef']))[0] : '';
    return cached
        || safeGet(ser, ['c:tx', 'c:strRef', 'c:f'])
        || safeGet(ser, ['c:tx', 'c:v'])
        || `\u7cfb\u5217${index + 1}`;
}

function getPtValue(point) {
    if (point == null) return '';
    if (typeof point !== 'object') return point;
    return point['c:v'] ?? '';
}

function getCachedPoints(refNode) {
    const points = safeGet(refNode, ['c:strCache', 'c:pt'])
        || safeGet(refNode, ['c:numCache', 'c:pt'])
        || safeGet(refNode, ['c:strLit', 'c:pt'])
        || safeGet(refNode, ['c:numLit', 'c:pt']);
    return sureArray(points).map(getPtValue);
}

function getMultiLevelCachedPoints(refNode) {
    const levels = sureArray(safeGet(refNode, ['c:multiLvlStrCache', 'c:lvl']));
    if (levels.length === 0) return [];

    const levelMaps = levels.map((level) => {
        const points = sureArray(level?.['c:pt']);
        const map = new Map();
        points.forEach((point, position) => {
            const rawIndex = Number(point?._$idx);
            const index = Number.isFinite(rawIndex) ? rawIndex : position;
            map.set(index, getPtValue(point));
        });
        return map;
    });
    const indexes = [...new Set(levelMaps.flatMap(map => [...map.keys()]))].sort((a, b) => a - b);

    return indexes.map(index => {
        const parts = levelMaps
            .map(map => map.get(index))
            .filter(value => !isBlankChartValue(value));
        return parts.join(' ');
    });
}

function getCachedCategoryPoints(refNode) {
    const cached = getCachedPoints(refNode);
    return cached.length ? cached : getMultiLevelCachedPoints(refNode);
}

function getCategoryData(ser, useCache = false) {
    const refNode = safeGet(ser, ['c:cat', 'c:strRef']) ||
        safeGet(ser, ['c:cat', 'c:numRef']) ||
        safeGet(ser, ['c:cat', 'c:multiLvlStrRef']);
    const catRef = refNode?.['c:f'];
    if (catRef) return catRef;

    const cached = useCache ? getCachedCategoryPoints(refNode) : [];
    if (cached.length > 0) return cached;

    return safeGet(ser, ['c:cat', 'c:strLit', 'c:pt'])
        || safeGet(ser, ['c:cat', 'c:numLit', 'c:pt'])
        || [];
}

// Extract category values.
const extractCategories = (ser, useCache = false) => {
    const catData = getCategoryData(ser, useCache);
    if (typeof catData === 'string') return catData // Return the address when a ref exists.
    if (Array.isArray(catData)) return catData.map(getPtValue); // Return inline literal values when present.
    return [] // No category data.
};

// Extract numeric values.
const extractValues = (ser) => {
    const valRef = safeGet(ser, ['c:val', 'c:numRef', 'c:f']);
    if (!valRef || typeof valRef !== 'string') return [];
    try {
        const valData = getDataByRef(valRef);
        if (Array.isArray(valData)) {
            return valData.map(v => {
                const num = parseFloat(v);
                return isNaN(num) ? null : num;
            });
        }
        const num = parseFloat(valData);
        return [isNaN(num) ? null : num];
    } catch (e) {
        return [];
    }
};

// Check whether a string is a chart ref with a sheet name.
export function isRefWithSheet(str) {
    if (typeof str !== 'string') return false;
    // ^ start
    // (?:'[^']+'|[^'!]+) sheet name, quoted or unquoted
    // ! separator
    // \$?[A-Za-z]{1,3}\$?\d+ a single cell ref, e.g. A1 or $B$2
    // (?::\$?[A-Za-z]{1,3}\$?\d+)? optional range end, e.g. :C3
    // $ end
    const regex = /^(?:'[^']+'|[^'!]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;
    return regex.test(str);
}

// Expand a range string into cell refs.
function rangeStrToCellRefs(rangeStr, SN) {
    if (typeof rangeStr !== 'string') return [];
    const normalized = resolveDefinedNameRef(rangeStr, SN);
    if (!normalized.includes('!')) return []; // A sheet name is required.
    let sheetName, cellRefs;
    [sheetName, cellRefs] = normalized.split('!');
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) sheetName = sheetName.slice(1, -1); // Strip outer single quotes.
    const arr = []
    SN.getSheet(sheetName).eachCells(cellRefs, (r, c) => {
        arr.push(sheetName + '!' + SN.Utils.cellNumToStr({ r, c }))
    })
    return arr
}

