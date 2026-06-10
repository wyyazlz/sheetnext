/**
 * Standard Chart XML Export Module
 * Provides build capabilities for OOXML standard diagrams
 */

import { isRefWithSheet } from "./chart.js";

const SERIES_EXT_URI = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}";
const SERIES_EXT_XMLNS = "http://schemas.microsoft.com/office/drawing/2014/chart";
const CHART_XML_ORDER = {
    chartSpace: ["c:date1904", "c:lang", "c:roundedCorners", "mc:AlternateContent", "c:style", "c:clrMapOvr", "c:pivotSource", "c:protection", "c:chart", "c:spPr", "c:txPr", "c:externalData", "c:printSettings", "c:userShapes", "c:extLst"],
    chart: ["c:title", "c:autoTitleDeleted", "c:pivotFmts", "c:view3D", "c:floor", "c:sideWall", "c:backWall", "c:plotArea", "c:legend", "c:plotVisOnly", "c:dispBlanksAs", "c:showDLblsOverMax", "c:extLst"],
    plotArea: ["c:layout", "c:areaChart", "c:area3DChart", "c:lineChart", "c:line3DChart", "c:stockChart", "c:radarChart", "c:scatterChart", "c:pieChart", "c:pie3DChart", "c:doughnutChart", "c:barChart", "c:bar3DChart", "c:ofPieChart", "c:surfaceChart", "c:surface3DChart", "c:bubbleChart", "c:valAx", "c:catAx", "c:dateAx", "c:serAx", "c:dTable", "c:spPr", "c:extLst"],
    title: ["c:tx", "c:layout", "c:overlay", "c:spPr", "c:txPr", "c:extLst"],
    legend: ["c:legendPos", "c:legendEntry", "c:layout", "c:overlay", "c:spPr", "c:txPr", "c:extLst"],
    dataLabels: ["c:dLbl", "c:delete", "c:numFmt", "c:spPr", "c:txPr", "c:dLblPos", "c:showLegendKey", "c:showVal", "c:showCatName", "c:showSerName", "c:showPercent", "c:showBubbleSize", "c:separator", "c:showLeaderLines", "c:leaderLines", "c:extLst"],
    dataLabel: ["c:idx", "c:delete", "c:layout", "c:tx", "c:numFmt", "c:spPr", "c:txPr", "c:dLblPos", "c:showLegendKey", "c:showVal", "c:showCatName", "c:showSerName", "c:showPercent", "c:showBubbleSize", "c:separator", "c:extLst"],
    manualLayout: ["c:layoutTarget", "c:xMode", "c:yMode", "c:wMode", "c:hMode", "c:x", "c:y", "c:w", "c:h", "c:extLst"],
    categoryAxis: ["c:axId", "c:scaling", "c:delete", "c:axPos", "c:majorGridlines", "c:minorGridlines", "c:title", "c:numFmt", "c:majorTickMark", "c:minorTickMark", "c:tickLblPos", "c:spPr", "c:txPr", "c:crossAx", "c:crosses", "c:crossesAt", "c:auto", "c:lblAlgn", "c:lblOffset", "c:tickLblSkip", "c:tickMarkSkip", "c:noMultiLvlLbl", "c:extLst"],
    valueAxis: ["c:axId", "c:scaling", "c:delete", "c:axPos", "c:majorGridlines", "c:minorGridlines", "c:title", "c:numFmt", "c:majorTickMark", "c:minorTickMark", "c:tickLblPos", "c:spPr", "c:txPr", "c:crossAx", "c:crosses", "c:crossesAt", "c:crossBetween", "c:majorUnit", "c:minorUnit", "c:dispUnits", "c:extLst"],
    barChart: ["c:barDir", "c:grouping", "c:varyColors", "c:ser", "c:dLbls", "c:gapWidth", "c:overlap", "c:serLines", "c:axId", "c:extLst"],
    lineChart: ["c:grouping", "c:varyColors", "c:ser", "c:dLbls", "c:dropLines", "c:hiLowLines", "c:upDownBars", "c:marker", "c:smooth", "c:axId", "c:extLst"],
    areaChart: ["c:grouping", "c:varyColors", "c:ser", "c:dLbls", "c:dropLines", "c:axId", "c:extLst"],
    scatterChart: ["c:scatterStyle", "c:varyColors", "c:ser", "c:dLbls", "c:axId", "c:extLst"],
    bubbleChart: ["c:varyColors", "c:ser", "c:dLbls", "c:bubble3D", "c:bubbleScale", "c:showNegBubbles", "c:sizeRepresents", "c:axId", "c:extLst"],
    radarChart: ["c:radarStyle", "c:varyColors", "c:ser", "c:dLbls", "c:axId", "c:extLst"],
    pieChart: ["c:varyColors", "c:ser", "c:dLbls", "c:firstSliceAng", "c:extLst"],
    doughnutChart: ["c:varyColors", "c:ser", "c:dLbls", "c:firstSliceAng", "c:holeSize", "c:extLst"],
    barSeries: ["c:idx", "c:order", "c:tx", "c:spPr", "c:invertIfNegative", "c:pictureOptions", "c:dPt", "c:dLbls", "c:trendline", "c:errBars", "c:cat", "c:val", "c:shape", "c:extLst"],
    lineSeries: ["c:idx", "c:order", "c:tx", "c:spPr", "c:marker", "c:pictureOptions", "c:dPt", "c:dLbls", "c:trendline", "c:errBars", "c:cat", "c:val", "c:smooth", "c:extLst"],
    areaSeries: ["c:idx", "c:order", "c:tx", "c:spPr", "c:pictureOptions", "c:dPt", "c:dLbls", "c:trendline", "c:errBars", "c:cat", "c:val", "c:extLst"],
    scatterSeries: ["c:idx", "c:order", "c:tx", "c:spPr", "c:marker", "c:dPt", "c:dLbls", "c:trendline", "c:errBars", "c:xVal", "c:yVal", "c:smooth", "c:extLst"],
    bubbleSeries: ["c:idx", "c:order", "c:tx", "c:spPr", "c:pictureOptions", "c:invertIfNegative", "c:dPt", "c:dLbls", "c:trendline", "c:errBars", "c:xVal", "c:yVal", "c:bubbleSize", "c:bubble3D", "c:extLst"],
    radarSeries: ["c:idx", "c:order", "c:tx", "c:spPr", "c:pictureOptions", "c:marker", "c:dPt", "c:dLbls", "c:cat", "c:val", "c:extLst"],
    pieSeries: ["c:idx", "c:order", "c:tx", "c:spPr", "c:pictureOptions", "c:explosion", "c:dPt", "c:dLbls", "c:cat", "c:val", "c:extLst"],
    textBody: ["a:bodyPr", "a:lstStyle", "a:p"],
    chartShapeProperties: ["a:xfrm", "a:custGeom", "a:prstGeom", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:ln", "a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst"],
    textCharacterProperties: ["a:ln", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill", "a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill", "a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"]
};
const CHART_GROUP_ORDER = {
    "c:barChart": CHART_XML_ORDER.barChart,
    "c:lineChart": CHART_XML_ORDER.lineChart,
    "c:areaChart": CHART_XML_ORDER.areaChart,
    "c:scatterChart": CHART_XML_ORDER.scatterChart,
    "c:bubbleChart": CHART_XML_ORDER.bubbleChart,
    "c:radarChart": CHART_XML_ORDER.radarChart,
    "c:pieChart": CHART_XML_ORDER.pieChart,
    "c:doughnutChart": CHART_XML_ORDER.doughnutChart
};
const CHART_SERIES_ORDER = {
    "c:barChart": CHART_XML_ORDER.barSeries,
    "c:lineChart": CHART_XML_ORDER.lineSeries,
    "c:areaChart": CHART_XML_ORDER.areaSeries,
    "c:scatterChart": CHART_XML_ORDER.scatterSeries,
    "c:bubbleChart": CHART_XML_ORDER.bubbleSeries,
    "c:radarChart": CHART_XML_ORDER.radarSeries,
    "c:pieChart": CHART_XML_ORDER.pieSeries,
    "c:doughnutChart": CHART_XML_ORDER.pieSeries
};

function sortXmlObjectInPlace(obj, keyOrder) {
    if (!obj || typeof obj !== 'object' || Array.isArray(obj) || !Array.isArray(keyOrder)) return;
    const ordered = {};
    keyOrder.forEach((key) => {
        if (Object.prototype.hasOwnProperty.call(obj, key)) ordered[key] = obj[key];
    });
    Object.keys(obj).forEach((key) => {
        if (!Object.prototype.hasOwnProperty.call(ordered, key)) ordered[key] = obj[key];
    });
    Object.keys(obj).forEach((key) => delete obj[key]);
    Object.keys(ordered).forEach((key) => {
        obj[key] = ordered[key];
    });
}

function normalizeTextCharacterProperties(node) {
    if (!node || typeof node !== 'object') return;
    sortXmlObjectInPlace(node, CHART_XML_ORDER.textCharacterProperties);
}

function normalizeShapeProperties(node) {
    if (!node || typeof node !== 'object') return;
    sortXmlObjectInPlace(node, CHART_XML_ORDER.chartShapeProperties);
}

function normalizeParagraphNode(node) {
    if (!node || typeof node !== 'object') return;
    if (node["a:pPr"]?.["a:defRPr"]) normalizeTextCharacterProperties(node["a:pPr"]["a:defRPr"]);
    const runs = Array.isArray(node["a:r"]) ? node["a:r"] : (node["a:r"] ? [node["a:r"]] : []);
    runs.forEach((run) => {
        if (run?.["a:rPr"]) normalizeTextCharacterProperties(run["a:rPr"]);
    });
    if (node["a:endParaRPr"]) normalizeTextCharacterProperties(node["a:endParaRPr"]);
}

function normalizeTextContainer(node) {
    if (!node || typeof node !== 'object') return;
    const paragraphs = Array.isArray(node["a:p"]) ? node["a:p"] : (node["a:p"] ? [node["a:p"]] : []);
    paragraphs.forEach(normalizeParagraphNode);
    sortXmlObjectInPlace(node, CHART_XML_ORDER.textBody);
}

function normalizeTextPropertiesNode(node) {
    if (!node || typeof node !== 'object') return;
    normalizeTextContainer(node);
}

function normalizeDataLabelNode(node) {
    if (!node || typeof node !== 'object') return;
    if (node["c:layout"]?.["c:manualLayout"]) {
        sortXmlObjectInPlace(node["c:layout"]["c:manualLayout"], CHART_XML_ORDER.manualLayout);
    }
    normalizeTextContainer(node["c:tx"]?.["c:rich"]);
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    if (node["c:txPr"]) normalizeTextPropertiesNode(node["c:txPr"]);
    sortXmlObjectInPlace(node, CHART_XML_ORDER.dataLabel);
}

function normalizeDataLabelsNode(node) {
    if (!node || typeof node !== 'object') return;
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    if (node["c:txPr"]) normalizeTextPropertiesNode(node["c:txPr"]);
    const labelNodes = Array.isArray(node["c:dLbl"]) ? node["c:dLbl"] : (node["c:dLbl"] ? [node["c:dLbl"]] : []);
    labelNodes.forEach(normalizeDataLabelNode);
    sortXmlObjectInPlace(node, CHART_XML_ORDER.dataLabels);
}

function normalizeTitleNode(node) {
    if (!node || typeof node !== 'object') return;
    normalizeTextContainer(node["c:tx"]?.["c:rich"]);
    if (node["c:layout"]?.["c:manualLayout"]) {
        sortXmlObjectInPlace(node["c:layout"]["c:manualLayout"], CHART_XML_ORDER.manualLayout);
    }
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    if (node["c:txPr"]) normalizeTextPropertiesNode(node["c:txPr"]);
    sortXmlObjectInPlace(node, CHART_XML_ORDER.title);
}

function normalizeLegendNode(node) {
    if (!node || typeof node !== 'object') return;
    if (node["c:layout"]?.["c:manualLayout"]) {
        sortXmlObjectInPlace(node["c:layout"]["c:manualLayout"], CHART_XML_ORDER.manualLayout);
    }
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    if (node["c:txPr"]) normalizeTextPropertiesNode(node["c:txPr"]);
    sortXmlObjectInPlace(node, CHART_XML_ORDER.legend);
}

function normalizeAxisNode(node, order) {
    if (!node || typeof node !== 'object') return;
    if (node["c:title"]) normalizeTitleNode(node["c:title"]);
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    if (node["c:txPr"]) normalizeTextPropertiesNode(node["c:txPr"]);
    sortXmlObjectInPlace(node, order);
}

function normalizeChartSeriesNode(node, chartKey) {
    if (!node || typeof node !== 'object') return;
    normalizeTextContainer(node["c:tx"]?.["c:rich"]);
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    if (node["c:dLbls"]) normalizeDataLabelsNode(node["c:dLbls"]);
    sortXmlObjectInPlace(node, CHART_SERIES_ORDER[chartKey]);
}

function normalizeChartGroupNode(node, chartKey) {
    if (!node || typeof node !== 'object') return;
    const seriesNodes = Array.isArray(node["c:ser"]) ? node["c:ser"] : (node["c:ser"] ? [node["c:ser"]] : []);
    seriesNodes.forEach((seriesNode) => normalizeChartSeriesNode(seriesNode, chartKey));
    if (node["c:dLbls"]) normalizeDataLabelsNode(node["c:dLbls"]);
    sortXmlObjectInPlace(node, CHART_GROUP_ORDER[chartKey]);
}

function normalizePlotAreaNode(node) {
    if (!node || typeof node !== 'object') return;
    if (node["c:layout"]?.["c:manualLayout"]) {
        sortXmlObjectInPlace(node["c:layout"]["c:manualLayout"], CHART_XML_ORDER.manualLayout);
    }
    ["c:areaChart", "c:lineChart", "c:radarChart", "c:scatterChart", "c:pieChart", "c:doughnutChart", "c:barChart", "c:bubbleChart"].forEach((chartKey) => {
        const chartNodes = Array.isArray(node[chartKey]) ? node[chartKey] : (node[chartKey] ? [node[chartKey]] : []);
        chartNodes.forEach((chartNode) => normalizeChartGroupNode(chartNode, chartKey));
    });
    const valueAxes = Array.isArray(node["c:valAx"]) ? node["c:valAx"] : (node["c:valAx"] ? [node["c:valAx"]] : []);
    valueAxes.forEach((axisNode) => normalizeAxisNode(axisNode, CHART_XML_ORDER.valueAxis));
    const categoryAxes = Array.isArray(node["c:catAx"]) ? node["c:catAx"] : (node["c:catAx"] ? [node["c:catAx"]] : []);
    categoryAxes.forEach((axisNode) => normalizeAxisNode(axisNode, CHART_XML_ORDER.categoryAxis));
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    sortXmlObjectInPlace(node, CHART_XML_ORDER.plotArea);
}

function normalizeChartSpaceNode(node) {
    if (!node || typeof node !== 'object') return;
    if (node["c:chart"]?.["c:title"]) normalizeTitleNode(node["c:chart"]["c:title"]);
    if (node["c:chart"]?.["c:plotArea"]) normalizePlotAreaNode(node["c:chart"]["c:plotArea"]);
    if (node["c:chart"]?.["c:legend"]) normalizeLegendNode(node["c:chart"]["c:legend"]);
    if (node["c:chart"]) sortXmlObjectInPlace(node["c:chart"], CHART_XML_ORDER.chart);
    if (node["c:spPr"]) normalizeShapeProperties(node["c:spPr"]);
    if (node["c:txPr"]) normalizeTextPropertiesNode(node["c:txPr"]);
    sortXmlObjectInPlace(node, CHART_XML_ORDER.chartSpace);
}

function toNumber(value, fallback = 0) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
}

function normalizeHexColor(value) {
    if (typeof value !== 'string') return null;
    const color = value.trim();
    if (!color) return null;
    const hex = color.startsWith('#') ? color.slice(1) : color;
    if (/^[0-9a-fA-F]{3}$/.test(hex)) {
        return hex.split('').map(char => char + char).join('').toUpperCase();
    }
    return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : null;
}

function clampAlpha(value, fallback = 1) {
    const num = Number(value);
    if (!Number.isFinite(num)) return fallback;
    return Math.max(0, Math.min(1, num));
}

function parseColorString(color) {
    if (typeof color !== 'string') return null;
    const text = color.trim();
    if (!text) return null;
    if (text.toLowerCase() === 'transparent') {
        return { hex: '000000', alpha: 0 };
    }
    const hex = normalizeHexColor(text);
    if (hex) return { hex, alpha: 1 };

    const rgbaMatch = text.match(/^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*([+-]?\d*\.?\d+))?\s*\)$/i);
    if (!rgbaMatch) return null;

    const toHex = (value) => {
        const num = Math.max(0, Math.min(255, Number(value)));
        return Math.round(num).toString(16).padStart(2, '0').toUpperCase();
    };

    return {
        hex: `${toHex(rgbaMatch[1])}${toHex(rgbaMatch[2])}${toHex(rgbaMatch[3])}`,
        alpha: rgbaMatch[4] == null ? 1 : clampAlpha(rgbaMatch[4], 1)
    };
}

function buildSolidFill(color) {
    const parsed = parseColorString(color);
    if (!parsed) return null;
    const colorNode = { "_$val": parsed.hex };
    if (parsed.alpha < 0.999) {
        colorNode["a:alpha"] = { "_$val": String(Math.round(parsed.alpha * 100000)) };
    }
    return { "a:srgbClr": colorNode };
}

function cloneChartXmlNode(node) {
    if (!node || typeof node !== 'object') return node;
    return JSON.parse(JSON.stringify(node));
}

function getSeriesStrokeColor(series) {
    return series?.lineStyle?.color ?? series?.itemStyle?.borderColor ?? series?.itemStyle?.color ?? null;
}

function getSeriesFillColor(series) {
    return series?.areaStyle?.color ?? series?.itemStyle?.color ?? null;
}

function getLineDashValue(type) {
    if (type === 'dashed') return 'dash';
    if (type === 'dotted') return 'sysDot';
    return null;
}

function buildShapeProperties({ fillColor = null, lineColor = null, lineWidthPx = null, lineDash = null, noLine = false, lineWidthAllowZero = false } = {}) {
    const spPr = {};
    const solidFill = buildSolidFill(fillColor);
    if (solidFill) spPr["a:solidFill"] = solidFill;

    const ln = {};
    const width = Number(lineWidthPx);
    if (Number.isFinite(width) && (width > 0 || (lineWidthAllowZero && width === 0))) {
        ln["_$w"] = width === 0 ? "0" : Math.max(1, Math.round(width * 9525));
    }
    const lineFill = buildSolidFill(lineColor);
    if (lineFill) {
        ln["a:solidFill"] = lineFill;
    } else if (noLine) {
        ln["a:noFill"] = "";
    }
    const dashValue = getLineDashValue(lineDash);
    if (dashValue) ln["a:prstDash"] = { "_$val": dashValue };
    if (Object.keys(ln).length > 0) spPr["a:ln"] = ln;

    return Object.keys(spPr).length > 0 ? { "c:spPr": spPr } : {};
}

function buildNoFillSpPr(includeEffectList = false) {
    const spPr = {
        "a:noFill": "",
        "a:ln": { "a:noFill": "" }
    };
    if (includeEffectList) spPr["a:effectLst"] = "";
    return spPr;
}

function buildThemeAxisLine() {
    return {
        "a:solidFill": {
            "a:schemeClr": {
                "a:lumMod": { "_$val": "15000" },
                "a:lumOff": { "_$val": "85000" },
                "_$val": "tx1"
            }
        },
        "a:round": "",
        "_$w": "9525",
        "_$cap": "flat",
        "_$cmpd": "sng",
        "_$algn": "ctr"
    };
}

function buildLineSpPrFromStyle(lineStyle = {}, fallback = null) {
    const shape = buildShapeProperties({
        lineColor: lineStyle?.color ?? null,
        lineWidthPx: lineStyle?.width,
        lineDash: lineStyle?.type
    })?.["c:spPr"];
    if (shape?.["a:ln"]) return shape;
    return fallback;
}

function isPlainWhiteColor(color) {
    const parsed = parseColorString(color);
    return !!parsed && parsed.alpha >= 0.999 && parsed.hex === 'FFFFFF';
}

function buildDefaultChartSpaceSpPr() {
    return {
        "a:solidFill": { "a:schemeClr": { "_$val": "bg1" } },
        "a:ln": {
            "a:solidFill": {
                "a:schemeClr": {
                    "a:lumMod": { "_$val": "15000" },
                    "a:lumOff": { "_$val": "85000" },
                    "_$val": "tx1"
                }
            },
            "a:round": "",
            "_$w": "9525",
            "_$cap": "flat",
            "_$cmpd": "sng",
            "_$algn": "ctr"
        },
        "a:effectLst": ""
    };
}

function buildLineNodeFromStyle(lineStyle = null, fallbackLine = null) {
    if (lineStyle?.show === false) {
        return { "a:noFill": "" };
    }
    const fallbackSpPr = fallbackLine ? { "a:ln": cloneChartXmlNode(fallbackLine) } : null;
    const spPr = buildLineSpPrFromStyle(lineStyle || {}, fallbackSpPr);
    if (spPr?.["a:ln"]) return spPr["a:ln"];
    return fallbackLine ? cloneChartXmlNode(fallbackLine) : { "a:noFill": "" };
}

function buildAreaSpPr(fillColor, lineStyle = null, fallbackSpPr = null, defaultSpPr = null) {
    const baseSpPr = fallbackSpPr && typeof fallbackSpPr === 'object' ? fallbackSpPr : defaultSpPr;
    const spPr = {};
    const fillNode = buildSolidFill(fillColor);
    if (fillNode) {
        spPr["a:solidFill"] = fillNode;
    } else if (baseSpPr?.["a:solidFill"]) {
        spPr["a:solidFill"] = cloneChartXmlNode(baseSpPr["a:solidFill"]);
    } else if (baseSpPr?.["a:noFill"] !== undefined) {
        spPr["a:noFill"] = "";
    }
    spPr["a:ln"] = buildLineNodeFromStyle(lineStyle, baseSpPr?.["a:ln"] || null);
    if (baseSpPr?.["a:effectLst"] !== undefined) {
        spPr["a:effectLst"] = cloneChartXmlNode(baseSpPr["a:effectLst"]);
    } else {
        spPr["a:effectLst"] = "";
    }
    return spPr;
}

function buildChartSpaceSpPr(chartOption = {}) {
    const defaultSpPr = buildDefaultChartSpaceSpPr();
    const fallbackSpPr = chartOption?.__excelChartAreaSpPr;
    const lineStyle = chartOption?.__excelChartAreaBorder ?? null;
    const hasRawSpPr = fallbackSpPr && typeof fallbackSpPr === 'object';
    const hasCustomFill = typeof chartOption?.backgroundColor === 'string' && !isPlainWhiteColor(chartOption.backgroundColor);
    if (!hasCustomFill && !lineStyle && !hasRawSpPr) {
        return defaultSpPr;
    }
    return buildAreaSpPr(
        hasCustomFill ? chartOption.backgroundColor : null,
        lineStyle,
        hasRawSpPr ? fallbackSpPr : null,
        defaultSpPr
    );
}

function buildPlotAreaSpPr(chartOption = {}) {
    const grid = chartOption?.grid || {};
    const hasGridStyle = typeof grid?.backgroundColor === 'string'
        || grid?.borderColor != null
        || grid?.borderWidth != null
        || grid?.borderType != null;
    const fallbackSpPr = chartOption?.__excelPlotAreaSpPr;
    if (!hasGridStyle) {
        return fallbackSpPr && typeof fallbackSpPr === 'object' ? cloneChartXmlNode(fallbackSpPr) : null;
    }
    const lineStyle = {
        show: true,
        ...(grid?.borderColor ? { color: grid.borderColor } : {}),
        ...(grid?.borderWidth != null ? { width: grid.borderWidth } : {}),
        ...(grid?.borderType ? { type: grid.borderType } : {})
    };
    return buildAreaSpPr(
        grid?.backgroundColor ?? null,
        lineStyle,
        fallbackSpPr && typeof fallbackSpPr === 'object' ? fallbackSpPr : null,
        null
    );
}

function getBackgroundContrastTextColor(backgroundColor) {
    const parsed = parseColorString(backgroundColor);
    if (!parsed) return null;
    const r = parseInt(parsed.hex.slice(0, 2), 16);
    const g = parseInt(parsed.hex.slice(2, 4), 16);
    const b = parseInt(parsed.hex.slice(4, 6), 16);
    const luminance = 0.299 * r + 0.587 * g + 0.114 * b;
    return luminance < 160 ? '#FFFFFF' : '#000000';
}

function getChartSpaceTextStyle(chartOption = {}) {
    const rawStyle = chartOption?.__excelChartSpaceTextStyle;
    if (rawStyle && typeof rawStyle === 'object' && Object.keys(rawStyle).length > 0) {
        return rawStyle;
    }
    const contrastColor = getBackgroundContrastTextColor(chartOption?.backgroundColor);
    return contrastColor ? { color: contrastColor } : {};
}

function buildChartSpaceTextProperties(chartOption = {}) {
    if (chartOption?.__excelChartSpaceTxPr) {
        return cloneChartXmlNode(chartOption.__excelChartSpaceTxPr);
    }
    return buildPlainTextProperties(getChartSpaceTextStyle(chartOption), 13.33)?.["c:txPr"] || {
        "a:bodyPr": "",
        "a:lstStyle": "",
        "a:p": {
            "a:pPr": { "a:defRPr": "" },
            "a:endParaRPr": { "_$lang": "zh-CN" }
        }
    };
}

function getMarkerSymbolValue(symbol) {
    switch (symbol) {
        case 'diamond':
        case 'triangle':
        case 'star':
        case 'circle':
            return symbol;
        case 'rect':
        case 'roundRect':
            return 'square';
        default:
            return 'circle';
    }
}

function parseLayoutRatio(value) {
    if (typeof value === 'number' && Number.isFinite(value)) {
        return value >= 0 && value <= 1 ? value : null;
    }
    if (typeof value !== 'string') return null;
    const text = value.trim();
    if (!text) return null;
    if (text.endsWith('%')) {
        const num = Number(text.slice(0, -1));
        return Number.isFinite(num) ? num / 100 : null;
    }
    const num = Number(text);
    return Number.isFinite(num) && num >= 0 && num <= 1 ? num : null;
}

function buildManualLayoutNode(layoutOption) {
    const { left, top, width, height } = layoutOption && typeof layoutOption === 'object' ? layoutOption : {};
    const x = parseLayoutRatio(left);
    const y = parseLayoutRatio(top);
    const w = parseLayoutRatio(width);
    const h = parseLayoutRatio(height);
    if (x == null && y == null && w == null && h == null) return "";
    const layout = {
        "c:layoutTarget": { "_$val": "inner" },
        "c:xMode": { "_$val": "edge" },
        "c:yMode": { "_$val": "edge" }
    };
    if (w != null) layout["c:wMode"] = { "_$val": "factor" };
    if (h != null) layout["c:hMode"] = { "_$val": "factor" };
    if (x != null) layout["c:x"] = { "_$val": x };
    if (y != null) layout["c:y"] = { "_$val": y };
    if (w != null) layout["c:w"] = { "_$val": w };
    if (h != null) layout["c:h"] = { "_$val": h };
    sortXmlObjectInPlace(layout, CHART_XML_ORDER.manualLayout);
    return { "c:manualLayout": layout };
}

function buildTextParagraphProps(textStyle = {}, fallbackSize = 14) {
    const fontSize = Math.round((textStyle?.fontSize ?? fallbackSize) * 75);
    const bold = textStyle?.fontWeight === 'bold' || Number(textStyle?.fontWeight) >= 600;
    const italic = textStyle?.fontStyle === 'italic';
    const fill = buildSolidFill(textStyle?.color);
    const defRPr = {
        "a:latin": { "_$typeface": "+mn-lt" },
        "a:ea": { "_$typeface": "+mn-ea" },
        "a:cs": { "_$typeface": "+mn-cs" },
        "_$sz": fontSize,
        "_$b": bold ? "1" : "0",
        "_$i": italic ? "1" : "0",
        "_$u": "none",
        "_$strike": "noStrike",
        "_$kern": "1200",
        "_$baseline": "0"
    };
    const rPr = {
        "_$lang": "zh-CN",
        "_$altLang": "en-US",
        "_$sz": fontSize,
        "_$b": bold ? "1" : "0",
        "_$i": italic ? "1" : "0"
    };
    if (fill) {
        defRPr["a:solidFill"] = fill;
        rPr["a:solidFill"] = fill;
    }
    sortXmlObjectInPlace(defRPr, CHART_XML_ORDER.textCharacterProperties);
    sortXmlObjectInPlace(rPr, CHART_XML_ORDER.textCharacterProperties);
    return { fontSize, defRPr, rPr };
}

function buildTextProperties(textStyle = {}, fallbackSize = 14, rotation = "0") {
    const { defRPr } = buildTextParagraphProps(textStyle, fallbackSize);
    return {
        "c:txPr": {
            "a:bodyPr": {
                "_$rot": rotation,
                "_$spcFirstLastPara": "1",
                "_$vertOverflow": "ellipsis",
                "_$vert": "horz",
                "_$wrap": "square",
                "_$anchor": "ctr",
                "_$anchorCtr": "1"
            },
            "a:lstStyle": "",
            "a:p": {
                "a:pPr": { "a:defRPr": defRPr },
                "a:endParaRPr": { "_$lang": "zh-CN" }
            }
        }
    };
}

function buildPlainTextProperties(textStyle = {}, fallbackSize = 14) {
    const { defRPr } = buildTextParagraphProps(textStyle, fallbackSize);
    return {
        "c:txPr": {
            "a:bodyPr": "",
            "a:lstStyle": "",
            "a:p": {
                "a:pPr": { "a:defRPr": defRPr },
                "a:endParaRPr": { "_$lang": "zh-CN" }
            }
        }
    };
}

function buildAxisTextProperties(axis, fallbackSize = 14) {
    if (axis?.__excelTransparentAxisLabel === true && axis?.__excelRawAxisTxPr) {
        return { "c:txPr": cloneChartXmlNode(axis.__excelRawAxisTxPr) };
    }
    return buildPlainTextProperties(axis?.axisLabel, fallbackSize);
}

function createGuid() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) {
        return crypto.randomUUID().toUpperCase();
    }
    const s4 = () => Math.floor((1 + Math.random()) * 0x10000).toString(16).padStart(4, '0').toUpperCase();
    return `${s4()}${s4()}-${s4()}-${s4()}-${s4()}-${s4()}${s4()}${s4()}`;
}

function buildSeriesUniqueId(index, baseGuid) {
    const normalized = String(baseGuid || '').replace(/[{}]/g, '').toUpperCase();
    const suffix = normalized.split('-').slice(1).join('-') || '0000-0000-0000-000000000000';
    const prefix = Number(index).toString(16).padStart(8, '0').toUpperCase();
    return `{${prefix}-${suffix}}`;
}

function buildSeriesExt(index, baseGuid) {
    return {
        "c:ext": {
            "_$uri": SERIES_EXT_URI,
            "_$xmlns:c16": SERIES_EXT_XMLNS,
            "c16:uniqueId": { "_$val": buildSeriesUniqueId(index, baseGuid) }
        }
    };
}

function buildLineMarker(series) {
    const show = series?.showSymbol === true || (typeof series?.symbolSize === 'number' && series.symbolSize > 0 && series?.showSymbol !== false);
    if (!show) {
        return { "c:marker": { "c:symbol": { "_$val": "none" } } };
    }
    const sizePx = typeof series?.symbolSize === 'number' && series.symbolSize > 0 ? series.symbolSize : 7.5;
    const sizeVal = Math.max(2, Math.round(sizePx * 72 / 96));
    const markerSpPr = buildLineSpPrFromStyle(
        { color: getSeriesStrokeColor(series) ?? getSeriesFillColor(series), width: 1 },
        null
    );
    return {
        "c:marker": {
            "c:symbol": { "_$val": getMarkerSymbolValue(series?.symbol) },
            "c:size": { "_$val": sizeVal },
            ...(markerSpPr ? { "c:spPr": markerSpPr } : {})
        }
    };
}

function buildLineStyle(series) {
    return buildShapeProperties({
        fillColor: series?.areaStyle ? getSeriesFillColor(series) : null,
        lineColor: getSeriesStrokeColor(series),
        lineWidthPx: toNumber(series?.lineStyle?.width, 2),
        lineDash: series?.lineStyle?.type
    });
}

function resolveRadarStyle(seriesList) {
    const radarSeries = seriesList.filter(s => s?.type === 'radar');
    if (radarSeries.some(s => s?.areaStyle)) return 'filled';
    const hasMarker = radarSeries.some(s => s?.showSymbol === true || (typeof s?.symbolSize === 'number' && s.symbolSize > 0 && s?.showSymbol !== false));
    return hasMarker ? 'marker' : 'standard';
}

function buildSeriesDataLabels(series) {
    const label = series?.label || {};
    const parts = label?.__excelParts || {};
    const formatter = typeof label?.formatter === 'string' ? label.formatter : '';
    const show = label?.show === true;
    const labels = getDefaultDLbls();
    const showSerName = parts.showSeriesName === true || formatter.includes('{a}');
    const showCatName = parts.showCategoryName === true || formatter.includes('{b}');
    const showVal = parts.showValue === true || (!showSerName && !showCatName && !parts.showPercent && !parts.showBubbleSize && show) || formatter.includes('{c}');
    const showPercent = parts.showPercent === true || formatter.includes('{d}');
    const showBubbleSize = parts.showBubbleSize === true;
    labels["c:showSerName"] = { "_$val": showSerName ? "1" : "0" };
    labels["c:showCatName"] = { "_$val": showCatName ? "1" : "0" };
    labels["c:showVal"] = { "_$val": showVal ? "1" : "0" };
    labels["c:showPercent"] = { "_$val": showPercent ? "1" : "0" };
    labels["c:showBubbleSize"] = { "_$val": showBubbleSize ? "1" : "0" };
    labels["c:numFmt"] = { "_$formatCode": "General", "_$sourceLinked": "0" };
    labels["c:spPr"] = buildNoFillSpPr(false);
    Object.assign(labels, buildPlainTextProperties(label, 12));
    if (typeof label?.__excelPosition === 'string' && label.__excelPosition) {
        labels["c:dLblPos"] = { "_$val": label.__excelPosition };
    } else if (show) {
        const position = label?.position;
        let excelPos = null;
        if (position === 'top') excelPos = 't';
        else if (position === 'bottom') excelPos = 'b';
        else if (position === 'left') excelPos = 'l';
        else if (position === 'right') excelPos = 'r';
        else if (position === 'inside') excelPos = 'ctr';
        else if (position === 'insideTop') excelPos = 'inEnd';
        else if (position === 'insideBottom') excelPos = 'inBase';
        else if (position === 'outside') excelPos = 'outEnd';
        if (excelPos) labels["c:dLblPos"] = { "_$val": excelPos };
    }
    labels["c:separator"] = ", ";
    labels["c:showLeaderLines"] = { "_$val": series?.labelLine?.show === true ? "1" : "0" };
    if (series?.type === 'pie' || Array.isArray(series?.radius)) {
        labels["c:showLeaderLines"] = { "_$val": series?.labelLine?.show === true ? "1" : "0" };
    }
    return { "c:dLbls": labels };
}

function getScatterStyle(series) {
    if (series?.__lineLike) {
        if (series?.smooth) return series?.showSymbol === false ? 'smooth' : 'smoothMarker';
        return series?.showSymbol === false ? 'line' : 'lineMarker';
    }
    return 'marker';
}

function buildScatterMarker(series) {
    const show = series?.showSymbol !== false;
    if (!show) return { "c:marker": { "c:symbol": { "_$val": "none" } } };
    const sizePx = typeof series?.symbolSize === 'number' && series.symbolSize > 0 ? series.symbolSize : 7.5;
    const sizeVal = Math.max(2, Math.round(sizePx * 72 / 96));
    const markerSpPr = buildShapeProperties({
        fillColor: getSeriesFillColor(series),
        lineColor: getSeriesStrokeColor(series) ?? getSeriesFillColor(series),
        lineWidthPx: 1
    })?.["c:spPr"];
    return {
        "c:marker": {
            "c:symbol": { "_$val": getMarkerSymbolValue(series?.symbol) },
            "c:size": { "_$val": sizeVal },
            ...(markerSpPr ? { "c:spPr": markerSpPr } : {})
        }
    };
}

function buildScatterLineStyle(style, series) {
    return buildShapeProperties({
        fillColor: getSeriesFillColor(series),
        lineColor: style === 'marker' ? null : getSeriesStrokeColor(series),
        lineWidthPx: style === 'marker' ? null : toNumber(series?.lineStyle?.width, 2),
        lineDash: style === 'marker' ? null : series?.lineStyle?.type,
        noLine: style === 'marker'
    });
}

function buildBarStyle(series) {
    const borderColor = series?.itemStyle?.borderColor ?? series?.lineStyle?.color ?? null;
    const borderWidth = series?.itemStyle?.borderWidth ?? series?.lineStyle?.width;
    const borderType = series?.itemStyle?.borderType ?? series?.lineStyle?.type;
    return buildShapeProperties({
        fillColor: getSeriesFillColor(series),
        lineColor: borderColor,
        lineWidthPx: borderColor || borderWidth != null || borderType ? borderWidth : 0,
        lineDash: borderType,
        noLine: !borderColor,
        lineWidthAllowZero: true
    });
}

function buildBubbleStyle(series) {
    return buildShapeProperties({
        fillColor: getSeriesFillColor(series),
        lineColor: getSeriesStrokeColor(series),
        lineWidthPx: series?.itemStyle?.borderWidth ?? series?.lineStyle?.width,
        lineDash: series?.lineStyle?.type
    });
}

function buildDataPointStyle(point) {
    return buildShapeProperties({
        fillColor: point?.itemStyle?.color ?? point?.color ?? null,
        lineColor: point?.itemStyle?.borderColor ?? null,
        lineWidthPx: point?.itemStyle?.borderWidth,
        lineDash: point?.itemStyle?.borderType
    });
}

function buildScatterSmooth(style) {
    const smooth = style === 'smooth' || style === 'smoothMarker';
    return { "c:smooth": { "_$val": smooth ? "1" : "0" } };
}

function extractSeriesValuesForExport(data) {
    if (!Array.isArray(data)) return [];
    return data.map((item) => {
        if (typeof item === 'number') return item;
        if (typeof item === 'string') return toNumber(item, 0);
        if (Array.isArray(item)) {
            const value = item.length > 1 ? item[1] : item[0];
            return toNumber(value, 0);
        }
        if (item && typeof item === 'object') {
            if (Object.prototype.hasOwnProperty.call(item, 'value')) {
                return toNumber(item.value, 0);
            }
        }
        return 0;
    });
}

function extractBubbleSizes(data) {
    if (!Array.isArray(data)) return [];
    return data.map((item) => {
        if (Array.isArray(item)) {
            if (item.length > 2) return toNumber(item[2], 0);
            if (item.length > 1) return toNumber(item[1], 0);
            return toNumber(item[0], 0);
        }
        if (item && typeof item === 'object') {
            const value = item.value;
            if (Array.isArray(value)) {
                if (value.length > 2) return toNumber(value[2], 0);
                if (value.length > 1) return toNumber(value[1], 0);
                return toNumber(value[0], 0);
            }
            if (Object.prototype.hasOwnProperty.call(item, 'value')) {
                return toNumber(item.value, 0);
            }
        }
        return toNumber(item, 0);
    });
}

function flattenNameValueData(data) {
    if (!Array.isArray(data)) return [];
    const items = [];
    data.forEach((item, idx) => {
        if (!item) return;
        if (typeof item === 'number' || typeof item === 'string') {
            items.push({ name: `项${idx + 1}`, value: toNumber(item, 0) });
            return;
        }
        if (Array.isArray(item)) {
            const value = item.length > 1 ? item[1] : item[0];
            items.push({ name: `项${idx + 1}`, value: toNumber(value, 0) });
            return;
        }
        if (typeof item === 'object') {
            const name = item.name ?? `项${idx + 1}`;
            if (Object.prototype.hasOwnProperty.call(item, 'value')) {
                items.push({ name, value: toNumber(item.value, 0) });
                return;
            }
            if (Array.isArray(item.children)) {
                const childTotal = item.children.reduce((sum, child) => sum + toNumber(child?.value, 0), 0);
                items.push({ name, value: childTotal });
            }
        }
    });
    return items;
}

function deriveCategoriesFromSeries(series) {
    if (!series?.length) return [];
    const first = series[0];
    if (Array.isArray(first.data)) {
        if (first.type === 'pie') {
            return first.data.map((item, idx) => item?.name ?? `项${idx + 1}`);
        }
        return first.data.map((item, idx) => {
            if (item && typeof item === 'object' && !Array.isArray(item)) {
                return item.name ?? `类别${idx + 1}`;
            }
            return `类别${idx + 1}`;
        });
    }
    return [];
}

function flattenNameValueDataSafe(data) {
    if (!Array.isArray(data)) return [];
    const items = [];
    data.forEach((item, idx) => {
        if (!item) return;
        if (typeof item === 'number' || typeof item === 'string') {
            items.push({ name: `\u9879${idx + 1}`, value: toNumber(item, 0) });
            return;
        }
        if (Array.isArray(item)) {
            const value = item.length > 1 ? item[1] : item[0];
            items.push({ name: `\u9879${idx + 1}`, value: toNumber(value, 0) });
            return;
        }
        if (typeof item === 'object') {
            const name = item.name ?? `\u9879${idx + 1}`;
            if (Object.prototype.hasOwnProperty.call(item, 'value')) {
                items.push({ name, value: toNumber(item.value, 0) });
                return;
            }
            if (Array.isArray(item.children)) {
                const childTotal = item.children.reduce((sum, child) => sum + toNumber(child?.value, 0), 0);
                items.push({ name, value: childTotal });
            }
        }
    });
    return items;
}

function deriveCategoriesFromSeriesSafe(series) {
    if (!series?.length) return [];
    const first = series[0];
    if (!Array.isArray(first.data)) return [];
    if (first.type === 'pie') {
        return first.data.map((item, idx) => item?.name ?? `\u9879${idx + 1}`);
    }
    return first.data.map((item, idx) => {
        if (item && typeof item === 'object' && !Array.isArray(item)) {
            return item.name ?? `\u7c7b\u522b${idx + 1}`;
        }
        return `\u7c7b\u522b${idx + 1}`;
    });
}

function isPercentAxis(axis) {
    if (!axis) return false;
    const formatter = axis?.axisLabel?.formatter;
    if (typeof formatter === 'string' && formatter.includes('%')) return true;
    const max = axis?.max;
    return max === 100 || max === '100';
}

function isPercentStackedChart(chartOption, valueAxis) {
    return chartOption?.__percent === true || isPercentAxis(valueAxis);
}

function normalizeAxisCollection(axisOption, fallback = {}) {
    const list = Array.isArray(axisOption)
        ? axisOption
        : (axisOption && typeof axisOption === 'object' ? [axisOption] : []);
    if (list.length === 0) return [{ ...fallback }];
    return list.map(axis => (axis && typeof axis === 'object' ? axis : { ...fallback }));
}

function normalizeChartOptionForExport(option) {
    const opt = option || {};
    const rawSeries = Array.isArray(opt.series) ? opt.series : [];
    const xAxes = normalizeAxisCollection(opt.xAxis, {});
    const yAxes = normalizeAxisCollection(opt.yAxis, {});
    const xAxis = xAxes[0];
    const yAxis = yAxes[0];
    const normalizedSeries = [];

    rawSeries.forEach((s) => {
        if (!s || !s.type) return;
        const hasPairData = Array.isArray(s.data) && Array.isArray(s.data[0]);
        if (s.type === 'line' && hasPairData) {
            normalizedSeries.push({ ...s, type: 'scatter', __lineLike: true });
            return;
        }
        switch (s.type) {
            case 'bar':
            case 'line':
            case 'pie':
            case 'scatter': {
                const isBubble = s.__bubble === true || typeof s.symbolSize === 'function';
                normalizedSeries.push(isBubble ? { ...s, __bubble: true } : s);
                break;
            }
            case 'radar':
                normalizedSeries.push(s);
                break;
            case 'candlestick': {
                const values = Array.isArray(s.data)
                    ? s.data.map(item => Array.isArray(item) ? item[1] : item)
                    : [];
                normalizedSeries.push({ ...s, type: 'line', data: extractSeriesValuesForExport(values) });
                break;
            }
            case 'funnel':
            case 'treemap':
            case 'sunburst': {
                const data = flattenNameValueDataSafe(s.data);
                normalizedSeries.push({ ...s, type: 'pie', data });
                break;
            }
            case 'waterfall': {
                normalizedSeries.push({ ...s, type: 'bar', data: extractSeriesValuesForExport(s.data) });
                break;
            }
            default: {
                normalizedSeries.push({ ...s, type: 'bar', data: extractSeriesValuesForExport(s.data) });
                break;
            }
        }
    });

    const expandedSeries = normalizedSeries.flatMap((s) => {
        if (s?.type !== 'radar') return [s];
        if (!Array.isArray(s.data) || s.data.length === 0) return [s];
        const isRadarDatumObject = (item) => {
            return item && typeof item === 'object' && !Array.isArray(item) && Object.prototype.hasOwnProperty.call(item, 'value');
        };
        if (!s.data.every(isRadarDatumObject)) return [s];
        return s.data.map((item) => ({
            ...s,
            name: item.name ?? s.name,
            data: item.value
        }));
    });

    const categories = xAxis?.data || yAxis?.data || deriveCategoriesFromSeriesSafe(expandedSeries);
    let exportXAxis = xAxis;
    let exportYAxis = yAxis;
    if (!exportXAxis && !exportYAxis && categories.length > 0) {
        exportXAxis = { type: 'category', data: categories };
        exportYAxis = { type: 'value' };
    }

    return {
        chartOption: opt,
        series: expandedSeries,
        xAxes,
        yAxes,
        xAxis: exportXAxis || {},
        yAxis: exportYAxis || {},
        radar: opt.radar
    };
}

/**
 * Get default data label settings
 */
function getDefaultDLbls() {
    return {
        "c:showLegendKey": { "_$val": "0" }, "c:showVal": { "_$val": "0" },
        "c:showCatName": { "_$val": "0" }, "c:showSerName": { "_$val": "0" },
        "c:showPercent": { "_$val": "0" }, "c:showBubbleSize": { "_$val": "0" }
    };
}

function buildTextTitle(text, styleOrSize = 14, layoutOption = null, overlay = false) {
    if (!text) return null;
    const textStyle = typeof styleOrSize === 'number' ? { fontSize: styleOrSize } : (styleOrSize || {});
    const { defRPr, rPr } = buildTextParagraphProps(textStyle, 14);
    const layout = buildManualLayoutNode(layoutOption);
    return {
        "c:title": {
            "c:tx": {
                "c:rich": {
                    "a:bodyPr": {
                        "_$rot": "0",
                        "_$spcFirstLastPara": "1",
                        "_$vertOverflow": "ellipsis",
                        "_$vert": "horz",
                        "_$wrap": "square",
                        "_$anchor": "ctr",
                        "_$anchorCtr": "1"
                    },
                    "a:lstStyle": "",
                    "a:p": {
                        "a:pPr": {
                            "a:defRPr": defRPr
                        },
                        "a:r": {
                            "a:rPr": rPr,
                            "a:t": String(text)
                        }
                    }
                }
            },
            "c:layout": layout,
            "c:overlay": { "_$val": overlay ? "1" : "0" },
            "c:spPr": { "a:noFill": "", "a:ln": { "a:noFill": "" }, "a:effectLst": "" }
        }
    };
}

function buildAxisTitleLayout(position = 'left') {
    let x = null;
    let y = null;
    if (position === 'left') {
        x = 0.08836023001133;
        y = 0.053111793229236176;
    } else if (position === 'right') {
        x = 0.831209304209126;
        y = 0.058761510743360466;
    } else if (position === 'bottom') {
        x = 0.42;
        y = 0.88;
    } else if (position === 'top') {
        x = 0.42;
        y = 0.02;
    }
    if (x == null || y == null) return "";
    return {
        "c:manualLayout": {
            "c:xMode": { "_$val": "edge" },
            "c:yMode": { "_$val": "edge" },
            "c:x": { "_$val": x },
            "c:y": { "_$val": y }
        }
    };
}

function buildAxisTitle(text, styleOrSize = 14, position = 'left') {
    if (!text) return null;
    const textStyle = typeof styleOrSize === 'number' ? { fontSize: styleOrSize } : (styleOrSize || {});
    const { defRPr, rPr } = buildTextParagraphProps(textStyle, 14);
    return {
        "c:title": {
            "c:tx": {
                "c:rich": {
                    "a:bodyPr": {
                        "_$rot": "0",
                        "_$spcFirstLastPara": "1",
                        "_$vertOverflow": "ellipsis",
                        "_$vert": "horz",
                        "_$wrap": "square",
                        "_$anchor": "ctr",
                        "_$anchorCtr": "1"
                    },
                    "a:lstStyle": "",
                    "a:p": {
                        "a:pPr": {
                            "a:defRPr": defRPr
                        },
                        "a:r": {
                            "a:rPr": rPr,
                            "a:t": String(text)
                        }
                    }
                }
            },
            "c:layout": buildAxisTitleLayout(position),
            "c:overlay": { "_$val": "0" },
            "c:spPr": { "a:noFill": "", "a:ln": { "a:noFill": "" }, "a:effectLst": "" }
        }
    };
}

function appendPlotAreaItem(plotArea, key, item) {
    if (!plotArea[key]) {
        plotArea[key] = item;
        return item;
    }
    if (Array.isArray(plotArea[key])) {
        plotArea[key].push(item);
        return item;
    }
    plotArea[key] = [plotArea[key], item];
    return item;
}

function appendPlotAreaItems(plotArea, key, items) {
    if (!Array.isArray(items) || items.length === 0) return;
    if (!plotArea[key]) {
        plotArea[key] = items.length === 1 ? items[0] : items;
        return;
    }
    items.forEach(item => appendPlotAreaItem(plotArea, key, item));
}

function getAxisFormatCode(axis, fallback = 'General') {
    if (typeof axis?.__excelFormatCode === 'string' && axis.__excelFormatCode.trim()) {
        return axis.__excelFormatCode.trim();
    }
    if (isPercentAxis(axis)) return '0%';
    return fallback;
}

function getAxisSourceLinked(axis, fallback = '1') {
    if (axis?.__excelSourceLinked === '0' || axis?.__excelSourceLinked === '1') {
        return axis.__excelSourceLinked;
    }
    return fallback;
}

function clampAxisIndex(index, length) {
    const value = Number(index);
    if (!Number.isFinite(value)) return 0;
    if (value < 0) return 0;
    if (value >= length) return Math.max(0, length - 1);
    return Math.floor(value);
}

function getAxisScaling(axis) {
    const scaling = {
        "c:orientation": { "_$val": axis?.inverse ? "maxMin" : "minMax" }
    };
    const min = Number(axis?.min);
    const max = Number(axis?.max);
    const logBase = Number(axis?.logBase);
    if (Number.isFinite(min)) scaling["c:min"] = { "_$val": min };
    if (Number.isFinite(max)) scaling["c:max"] = { "_$val": max };
    if ((axis?.type === 'log' || Number.isFinite(logBase)) && Number.isFinite(logBase || 10)) {
        scaling["c:logBase"] = { "_$val": Number.isFinite(logBase) ? logBase : 10 };
    }
    return scaling;
}

function applyAxisScale(obj, axis, withGridlines = false, defaultCrosses = null) {
    obj["c:scaling"] = getAxisScaling(axis);
    const majorUnit = Number(axis?.interval);
    const minorUnit = Number(axis?.__excelMinorUnit);
    if (Number.isFinite(majorUnit)) {
        obj["c:majorUnit"] = { "_$val": majorUnit };
    }
    if (Number.isFinite(minorUnit)) {
        obj["c:minorUnit"] = { "_$val": minorUnit };
    }
    const crossesAt = Number(axis?.__excelCrossesAt ?? axis?.crossesAt);
    if (Number.isFinite(crossesAt)) {
        obj["c:crossesAt"] = { "_$val": crossesAt };
        delete obj["c:crosses"];
    } else {
        const crosses = axis?.__excelCrosses || defaultCrosses || (axis?.boundaryGap === false ? 'min' : 'autoZero');
        obj["c:crosses"] = { "_$val": crosses };
        delete obj["c:crossesAt"];
    }
    if (typeof axis?.__excelCrossBetween === 'string' && axis.__excelCrossBetween) {
        obj["c:crossBetween"] = { "_$val": axis.__excelCrossBetween };
    }
    if (withGridlines && axis?.splitLine?.show === false) {
        delete obj["c:majorGridlines"];
    }
    return obj;
}

function buildAxisSpPr(axis) {
    if (axis?.__excelTransparentAxisLabel === true && axis?.__excelRawAxisSpPr) {
        return cloneChartXmlNode(axis.__excelRawAxisSpPr);
    }
    if (axis?.__excelTransparentAxisLabel === true) {
        return buildNoFillSpPr(true);
    }
    if (axis?.axisLine?.show === false) {
        return buildNoFillSpPr(true);
    }
    const lineStyle = axis?.axisLine?.lineStyle || null;
    const spPr = buildLineSpPrFromStyle(lineStyle, { "a:ln": buildThemeAxisLine() }) || { "a:ln": buildThemeAxisLine() };
    spPr["a:effectLst"] = "";
    return spPr;
}

function buildGridlineNode(splitLine, defaultMode = 'theme') {
    if (splitLine?.show === false) return null;
    if (!splitLine && !defaultMode) return null;

    const fallbackLine = defaultMode === 'light'
        ? {
            "a:solidFill": { "a:srgbClr": { "_$val": "D9D9D9" } },
            "_$w": "9525"
        }
        : buildThemeAxisLine();

    const lineStyle = splitLine?.lineStyle || null;
    const spPr = buildLineSpPrFromStyle(lineStyle, { "a:ln": fallbackLine }) || { "a:ln": fallbackLine };
    spPr["a:effectLst"] = "";
    return {
        "c:majorGridlines": {
            "c:spPr": spPr
        }
    };
}

/**
 * Creating a histogram
 */
function createBarChart(id1, id2, isStack, barDir, isPercent = false) {
    return {
        "c:barDir": { "_$val": barDir },
        "c:grouping": { "_$val": isStack ? (isPercent ? 'percentStacked' : 'stacked') : "clustered" },
        "c:varyColors": { "_$val": "0" },
        "c:ser": [],
        "c:dLbls": getDefaultDLbls(),
        "c:gapWidth": { "_$val": "150" },
        "c:overlap": { "_$val": isStack ? "100" : "-27" },
        "c:axId": [{ "_$val": id1 }, { "_$val": id2 }]
    };
}

/**
 * Create Line Chart
 */
function createLineChart(id1, id2, grouping = 'standard') {
    return {
        "c:grouping": { "_$val": grouping },
        "c:varyColors": { "_$val": "0" },
        "c:ser": [],
        "c:dLbls": getDefaultDLbls(),
        "c:smooth": { "_$val": "0" },
        "c:axId": [{ "_$val": id1 }, { "_$val": id2 }]
    };
}

/**
 * Create Scatter Chart
 */
function createScatterChart(id1, id2, style = 'marker') {
    return {
        "c:scatterStyle": { "_$val": style },
        "c:varyColors": { "_$val": "0" },
        "c:ser": [],
        "c:axId": [{ "_$val": id1 }, { "_$val": id2 }]
    };
}

/**
 * Create Bubble Chart
 */
function createBubbleChart(id1, id2) {
    return {
        "c:varyColors": { "_$val": "0" },
        "c:ser": [],
        "c:dLbls": getDefaultDLbls(),
        "c:bubbleScale": { "_$val": "100" },
        "c:showNegBubbles": { "_$val": "0" },
        "c:sizeRepresents": { "_$val": "area" },
        "c:axId": [{ "_$val": id1 }, { "_$val": id2 }]
    };
}

/**
 * Create Radar Chart
 */
function createRadarChart(id1, id2, style = 'standard') {
    return {
        "c:radarStyle": { "_$val": style },
        "c:varyColors": { "_$val": "0" },
        "c:ser": [],
        "c:dLbls": getDefaultDLbls(),
        "c:axId": [{ "_$val": id1 }, { "_$val": id2 }]
    };
}

/**
 * Create pie chart
 */
function createPieChart(s) {
    const startAngle = Number(s?.startAngle);
    const obj = {
        "c:varyColors": { "_$val": "1" },
        "c:ser": {},
        "c:dLbls": { ...getDefaultDLbls(), "c:showLeaderLines": { "_$val": "0" } },
        "c:firstSliceAng": { "_$val": Number.isFinite(startAngle) ? String(Math.round(((startAngle % 360) + 360) % 360)) : "0" }
    };
    if (Array.isArray(s.radius)) {
        const inner = parseLayoutRatio(s.radius[0]);
        const outer = parseLayoutRatio(s.radius[1]);
        const holeSize = inner != null && outer != null && outer > 0
            ? Math.max(10, Math.min(90, Math.round((inner / outer) * 100)))
            : 75;
        obj['c:holeSize'] = { _$val: String(holeSize) };
    }
    return obj;
}

/**
 * Get taxonomy axis
 */
function getCatAx(id1, id2, axisOption, position = 'bottom', extra = {}) {
    const axis = axisOption || {};
    const hidden = extra?.hidden === true || axis?.show === false;
    const axisPos = position === 'top' ? 't' : (position === 'left' ? 'l' : (position === 'right' ? 'r' : 'b'));
    const majorGridlines = axis?.splitLine?.show ? buildGridlineNode(axis.splitLine, 'light') : null;
    const keepTransparentLabelSlot = axis?.__excelTransparentAxisLabel === true;
    const sourceLinked = getAxisSourceLinked(axis, hidden ? '1' : '0');
    const obj = {
        "c:axId": { "_$val": id1 },
        "c:scaling": getAxisScaling(axis),
        "c:delete": { "_$val": hidden ? "1" : "0" },
        "c:axPos": { "_$val": axisPos },
        ...(majorGridlines || {}),
        "c:numFmt": { "_$formatCode": "General", "_$sourceLinked": sourceLinked },
        "c:majorTickMark": { "_$val": "none" },
        "c:minorTickMark": { "_$val": "none" },
        "c:tickLblPos": { "_$val": axis?.axisLabel?.show === false && !keepTransparentLabelSlot ? "none" : "nextTo" },
        "c:spPr": buildAxisSpPr(axis),
        ...buildAxisTextProperties(axis, axis?.axisLabel?.fontSize ?? 12),
        "c:crossAx": { "_$val": id2 },
        "c:crosses": { "_$val": axis.boundaryGap === false ? "min" : "autoZero" },
        "c:auto": { "_$val": "1" },
        "c:lblAlgn": { "_$val": "ctr" },
        "c:lblOffset": { "_$val": "0" },
        "c:tickMarkSkip": { "_$val": "1" },
        "c:noMultiLvlLbl": { "_$val": "1" }
    };
    applyAxisScale(obj, axis);
    const title = buildAxisTitle(axis.name, axis.nameTextStyle ?? { fontSize: axis.axisLabel?.fontSize ?? 14 }, position);
    if (title && !hidden) Object.assign(obj, title);
    return obj;
}

/**
 * Get Value Axis
 */
function getValAx(id1, id2, axisOption, formatCode = 'General', position = 'left', extra = {}) {
    const axis = axisOption || {};
    const hidden = extra?.hidden === true || axis?.show === false;
    const axisPos = position === 'right' ? 'r' : (position === 'top' ? 't' : (position === 'bottom' ? 'b' : 'l'));
    const majorGridlines = buildGridlineNode(axis?.splitLine, 'theme');
    const keepTransparentLabelSlot = axis?.__excelTransparentAxisLabel === true;
    const sourceLinked = getAxisSourceLinked(axis, position === 'left' || position === 'bottom' ? '0' : '1');
    const obj = {
        "c:axId": { "_$val": id2 },
        "c:scaling": getAxisScaling(axis),
        "c:delete": { "_$val": hidden ? "1" : "0" },
        "c:axPos": { "_$val": axisPos },
        ...(majorGridlines || {}),
        "c:numFmt": { "_$formatCode": formatCode || "General", "_$sourceLinked": sourceLinked },
        "c:majorTickMark": { "_$val": "none" },
        "c:minorTickMark": { "_$val": "none" },
        "c:tickLblPos": { "_$val": axis?.axisLabel?.show === false && !keepTransparentLabelSlot ? "none" : "nextTo" },
        "c:spPr": buildAxisSpPr(axis),
        ...buildAxisTextProperties(axis, axis?.axisLabel?.fontSize ?? 12),
        "c:crossAx": { "_$val": id1 },
        "c:crosses": { "_$val": axis.boundaryGap === false ? "min" : "autoZero" },
        "c:crossBetween": { "_$val": "between" }
    };
    applyAxisScale(obj, axis, true, extra?.defaultCrosses || null);
    const title = buildAxisTitle(axis.name, axis.nameTextStyle ?? { fontSize: axis.axisLabel?.fontSize ?? 14 }, position);
    if (title && !hidden) Object.assign(obj, title);
    return obj;
}

/**
 * Get Scatterplot Value Axis
 */
function getScatterValAx(id, crossId, axis, pos) {
    const majorGridlines = buildGridlineNode(axis?.splitLine, 'theme');
    const obj = {
        "c:axId": { "_$val": id },
        "c:scaling": getAxisScaling(axis),
        "c:delete": { "_$val": "0" },
        "c:axPos": { "_$val": pos },
        ...(majorGridlines || {}),
        "c:numFmt": { "_$formatCode": getAxisFormatCode(axis, "General"), "_$sourceLinked": "1" },
        "c:majorTickMark": { "_$val": "none" },
        "c:minorTickMark": { "_$val": "none" },
        "c:tickLblPos": { "_$val": "nextTo" },
        "c:spPr": buildAxisSpPr(axis),
        ...buildAxisTextProperties(axis, axis?.axisLabel?.fontSize ?? 12),
        "c:crossAx": { "_$val": crossId },
        "c:crosses": { "_$val": axis?.boundaryGap === false ? "min" : "autoZero" },
        "c:crossBetween": { "_$val": "between" }
    };
    applyAxisScale(obj, axis, true);
    const title = buildAxisTitle(axis?.name, axis?.nameTextStyle ?? { fontSize: axis?.axisLabel?.fontSize ?? 14 }, pos === 'r' ? 'right' : (pos === 't' ? 'top' : (pos === 'b' ? 'bottom' : 'left')));
    if (title) Object.assign(obj, title);
    return obj;
}

/**
 * Building Data References
 */
function buildRef(str, type) {
    if (str == null) {
        if (type == 'str') return { "c:strLit": { "c:ptCount": { "_$val": 0 } } };
        return { "c:numLit": { "c:formatCode": "General", "c:ptCount": { "_$val": 0 } } };
    }
    if (isRefWithSheet(str)) {
        const obj = {};
        obj['c:' + (type == 'str' ? 'strRef' : 'numRef')] = { "c:f": str };
        return obj;
    } else if (Array.isArray(str)) {
        const key = 'c:' + (type == 'str' ? 'strLit' : 'numLit');
        if (str.length === 0) {
            if (type == 'str') return { [key]: { "c:ptCount": { "_$val": 0 } } };
            return { [key]: { "c:formatCode": "General", "c:ptCount": { "_$val": 0 } } };
        }
        const lit = type == 'str'
            ? {
                "c:ptCount": { "_$val": str.length },
                "c:pt": str.map((val, idx) => ({ "_$idx": idx, "c:v": val == null ? '' : String(val) }))
            }
            : {
                "c:formatCode": "General",
                "c:ptCount": { "_$val": str.length },
                "c:pt": str.map((val, idx) => ({ "_$idx": idx, "c:v": val == null ? '' : String(val) }))
            };
        return { [key]: lit };
    } else {
        return { "c:v": String(str) };
    }
}

function buildSeriesText(value) {
    if (value == null) return { "c:v": "" };
    if (isRefWithSheet(value)) {
        return { "c:strRef": { "c:f": value } };
    }
    return { "c:v": String(value) };
}

/**
 * Processing Cell References
 */
function processCellReferences(SN, refs) {
    return refs.map(r => {
        if (typeof r === 'string' && r.includes('!') && isRefWithSheet(r)) {
            try {
                const match = r.match(/^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/);
                if (match) {
                    const sheetName = (match[1] ?? match[2] ?? '').replace(/''/g, "'");
                    const cellRef = match[3].replace(/\$/g, '');
                    if (cellRef.includes(':')) return r;
                    const sheet = SN.getSheet(sheetName);
                    const cell = sheet.getCell(cellRef);
                    return cell.calcVal;
                }
            } catch (e) { }
        }
        return r;
    });
}

function buildChartTitleNode(titleOption = {}) {
    const titleText = titleOption?.text ?? '';
    if (!titleText || titleOption?.show === false) return null;
    return buildTextTitle(titleText, titleOption?.textStyle || {}, titleOption, !!titleOption?.overlay)?.["c:title"] || null;
}

function resolveLegendPosition(legend = {}) {
    if (legend?.left === 'left') return 'l';
    if (legend?.right === 'right') return 'r';
    if (legend?.top != null && legend?.left === 'center') return 't';
    if (legend?.top === 'middle' && legend?.right === 'right') return 'r';
    if (legend?.top === 'middle' && legend?.left === 'left') return 'l';
    return legend?.bottom != null ? 'b' : 't';
}

function buildLegendNode(legendOption = {}) {
    if (legendOption?.show === false) return null;
    const legendPos = resolveLegendPosition(legendOption);
    const layout = buildManualLayoutNode(legendOption);
    return {
        "c:legendPos": { "_$val": legendPos },
        "c:layout": layout,
        "c:overlay": { "_$val": legendOption?.overlay ? "1" : "0" },
        "c:spPr": { "a:noFill": "", "a:ln": { "a:noFill": "" }, "a:effectLst": "" },
        ...buildTextProperties(legendOption?.textStyle || {}, 14, "0")
    };
}

function buildPlotAreaLayout(grid = {}) {
    const left = parseLayoutRatio(grid?.left);
    const top = parseLayoutRatio(grid?.top);
    const right = parseLayoutRatio(grid?.right);
    const bottom = parseLayoutRatio(grid?.bottom);
    if (left == null || top == null || right == null || bottom == null) return "";
    const width = 1 - left - right;
    const height = 1 - top - bottom;
    if (width <= 0 || height <= 0) return "";
    return buildManualLayoutNode({ left, top, width, height });
}

/**
 * Creating a Chart Space Object
 */
function createChartSpace(d, plotArea) {
    const titleNode = buildChartTitleNode(d.chartOption?.title || {});
    const legendNode = buildLegendNode(d.chartOption?.legend || {});
    const chartSpaceSpPr = buildChartSpaceSpPr(d.chartOption || {});
    const chartSpaceTxPr = buildChartSpaceTextProperties(d.chartOption || {});
    const chartNode = {
        "c:autoTitleDeleted": { "_$val": titleNode ? "0" : "1" },
        "c:plotArea": plotArea,
        "c:plotVisOnly": { "_$val": "1" },
        "c:dispBlanksAs": { "_$val": "gap" },
        "c:showDLblsOverMax": { "_$val": "0" }
    };
    if (titleNode) chartNode["c:title"] = titleNode;
    if (legendNode) chartNode["c:legend"] = legendNode;
    return {
        "c:date1904": { "_$val": "0" },
        "c:lang": { "_$val": "zh-CN" },
        "c:roundedCorners": { "_$val": "0" },
        "mc:AlternateContent": { "mc:Choice": { "c14:style": { "_$val": "102" }, "_$Requires": "c14", "_$xmlns:c14": "http://schemas.microsoft.com/office/drawing/2007/8/2/chart" }, "mc:Fallback": { "c:style": { "_$val": "2" } }, "_$xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006" },
        "c:chart": chartNode,
        "c:spPr": chartSpaceSpPr,
        "c:txPr": chartSpaceTxPr,
        "c:printSettings": { "c:headerFooter": "", "c:pageMargins": { "_$b": "0.75", "_$l": "0.7", "_$r": "0.7", "_$t": "0.75", "_$header": "0.3", "_$footer": "0.3" }, "c:pageSetup": "" },
        "_$xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "_$xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "_$xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    };
}

/**
 * Building a Chart Graph Frame
 */
export function buildChartGraphicFrame(i, chartName) {
    return {
        "xdr:nvGraphicFramePr": { "xdr:cNvPr": { "_$id": i + 1, "_$name": `图表 ${i + 1}` }, "xdr:cNvGraphicFramePr": "" },
        "xdr:xfrm": { "a:off": { "_$x": "0", "_$y": "0" }, "a:ext": { "_$cx": "0", "_$cy": "0" } },
        "a:graphic": { "a:graphicData": { "c:chart": { "_$xmlns:c": "http://schemas.openxmlformats.org/drawingml/2006/chart", "_$xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "_$r:id": 'rId' + (i + 1) }, "_$uri": "http://schemas.openxmlformats.org/drawingml/2006/chart" } },
        "_$macro": ""
    };
}

/**
 * Build Picture Object
 */
export function buildImagePic(i, rotation) {
    const xfrm = { "a:off": { "_$x": "5381625", "_$y": "590550" }, "a:ext": { "_$cx": "1295238", "_$cy": "1009524" } };
    if (rotation) xfrm["_$rot"] = String(Math.round(rotation * 60000));
    return {
        "xdr:nvPicPr": { "xdr:cNvPr": { "_$id": i + 1, "_$name": "图片 " + (i + 1) }, "xdr:cNvPicPr": { "a:picLocks": { "_$noChangeAspect": "1" } } },
        "xdr:blipFill": { "a:blip": { "_$xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "_$r:embed": "rId" + (i + 1) }, "a:stretch": { "a:fillRect": "" } },
        "xdr:spPr": { "a:xfrm": xfrm, "a:prstGeom": { "a:avLst": "", "_$prst": "rect" } }
    };
}

/**
 * Get chart XML details
 * @param {Workbook} SN - Workbook Instances
 * @param {Drawing} d - Chart Drawing Objects
 * @ returns {Object} chart XML object
 */
export function getChartXmlDetail(SN, d) {
    const exportInfo = normalizeChartOptionForExport(d.chartOption);
    const chartOption = exportInfo.chartOption;
    const xAxes = exportInfo.xAxes || [exportInfo.xAxis || {}];
    const yAxes = exportInfo.yAxes || [exportInfo.yAxis || {}];
    const xAxis = exportInfo.xAxis || {};
    const yAxis = exportInfo.yAxis || {};
    const radar = exportInfo.radar;
    const seriesList = exportInfo.series;
    const plotArea = { "c:layout": buildPlotAreaLayout(chartOption?.grid) };
    const radarChartStyle = resolveRadarStyle(seriesList);
    let num = 1000000000;
    const nextAxisId = () => ++num;
    const chartGuid = createGuid();
    let hasChart = false;
    let categoryAxisSetup = null;
    const categoryChartGroups = new Map();
    const scatterAxisBindings = new Map();
    const scatterChartGroups = new Map();

    const toScatterAxisPos = (position, fallback) => {
        const pos = position || fallback;
        if (pos === 'top') return 't';
        if (pos === 'bottom') return 'b';
        if (pos === 'right') return 'r';
        return 'l';
    };

    const ensureCategoryAxisSetup = (barDir = 'col') => {
        if (categoryAxisSetup) return categoryAxisSetup;
        const categoryOnX = barDir !== 'bar';
        categoryAxisSetup = {
            categoryOnX,
            categoryAxes: normalizeAxisCollection(categoryOnX ? xAxes : yAxes, { type: 'category' }),
            valueAxes: normalizeAxisCollection(categoryOnX ? yAxes : xAxes, { type: 'value' }),
            bindings: new Map(),
            categoryReuseCounts: new Map(),
            valueReuseCounts: new Map(),
            categoryAxisNodes: [],
            valueAxisNodes: [],
            flushed: false
        };
        return categoryAxisSetup;
    };

    const resolveCategoryAxisBinding = (axisSetup, categoryIndex, valueIndex) => {
        const key = `${categoryIndex}|${valueIndex}`;
        if (axisSetup.bindings.has(key)) return axisSetup.bindings.get(key);

        const categoryAxis = axisSetup.categoryAxes[categoryIndex] || axisSetup.categoryAxes[0] || {};
        const valueAxis = axisSetup.valueAxes[valueIndex] || axisSetup.valueAxes[0] || {};
        const categoryReuseIndex = axisSetup.categoryReuseCounts.get(categoryIndex) || 0;
        const valueReuseIndex = axisSetup.valueReuseCounts.get(valueIndex) || 0;
        axisSetup.categoryReuseCounts.set(categoryIndex, categoryReuseIndex + 1);
        axisSetup.valueReuseCounts.set(valueIndex, valueReuseIndex + 1);

        const hiddenCategoryAxis = categoryReuseIndex > 0;
        const hiddenValueAxis = valueReuseIndex > 0;
        const categoryId = nextAxisId();
        const valueId = nextAxisId();
        const defaultCategoryPosition = axisSetup.categoryOnX
            ? (hiddenCategoryAxis ? 'bottom' : (categoryIndex === 0 ? 'bottom' : 'top'))
            : (hiddenCategoryAxis ? 'left' : (categoryIndex === 0 ? 'left' : 'right'));
        const defaultValuePosition = axisSetup.categoryOnX
            ? (valueIndex === 0 ? 'left' : 'right')
            : (valueIndex === 0 ? 'bottom' : 'top');
        const valuePosition = valueAxis.position || defaultValuePosition;

        axisSetup.categoryAxisNodes.push(getCatAx(
            categoryId,
            valueId,
            categoryAxis,
            categoryAxis.position || defaultCategoryPosition,
            { hidden: hiddenCategoryAxis }
        ));
        axisSetup.valueAxisNodes.push(getValAx(
            categoryId,
            valueId,
            valueAxis,
            getAxisFormatCode(valueAxis, 'General'),
            valuePosition,
            {
                hidden: hiddenValueAxis,
                defaultCrosses: (valuePosition === 'right' || valuePosition === 'top') ? 'max' : null
            }
        ));

        const binding = {
            categoryIndex,
            valueIndex,
            categoryId,
            valueId,
            categories: categoryAxis?.data ?? [],
            valueAxis
        };
        axisSetup.bindings.set(key, binding);
        return binding;
    };

    const flushCategoryAxes = () => {
        if (!categoryAxisSetup || categoryAxisSetup.flushed) return;
        appendPlotAreaItems(plotArea, "c:catAx", categoryAxisSetup.categoryAxisNodes);
        appendPlotAreaItems(plotArea, "c:valAx", categoryAxisSetup.valueAxisNodes);
        categoryAxisSetup.flushed = true;
    };

    const getCategoryAxisBinding = (series, barDir = 'col') => {
        const axisSetup = ensureCategoryAxisSetup(barDir);
        const categoryIndex = clampAxisIndex(
            axisSetup.categoryOnX ? series?.xAxisIndex : series?.yAxisIndex,
            axisSetup.categoryAxes.length
        );
        const valueIndex = clampAxisIndex(
            axisSetup.categoryOnX ? series?.yAxisIndex : series?.xAxisIndex,
            axisSetup.valueAxes.length
        );
        const binding = resolveCategoryAxisBinding(axisSetup, categoryIndex, valueIndex);
        return {
            axisSetup,
            ...binding
        };
    };

    const getChartGroup = (propKey, signature, factory) => {
        const key = `${propKey}|${signature}`;
        if (categoryChartGroups.has(key)) return categoryChartGroups.get(key);
        const group = factory();
        appendPlotAreaItem(plotArea, propKey, group);
        categoryChartGroups.set(key, group);
        return group;
    };

    const getScatterAxisBinding = (series) => {
        const xIndex = clampAxisIndex(series?.xAxisIndex, xAxes.length);
        const yIndex = clampAxisIndex(series?.yAxisIndex, yAxes.length);
        const key = `${xIndex}|${yIndex}`;
        if (scatterAxisBindings.has(key)) return scatterAxisBindings.get(key);

        const scatterXAxis = xAxes[xIndex] || xAxis || {};
        const scatterYAxis = yAxes[yIndex] || yAxis || {};
        const xId = nextAxisId();
        const yId = nextAxisId();

        appendPlotAreaItem(
            plotArea,
            "c:valAx",
            getScatterValAx(xId, yId, scatterXAxis, toScatterAxisPos(scatterXAxis.position, xIndex === 0 ? 'bottom' : 'top'))
        );
        appendPlotAreaItem(
            plotArea,
            "c:valAx",
            getScatterValAx(yId, xId, scatterYAxis, toScatterAxisPos(scatterYAxis.position, yIndex === 0 ? 'left' : 'right'))
        );

        const binding = { key, xId, yId };
        scatterAxisBindings.set(key, binding);
        return binding;
    };

    const getScatterChartGroup = (propKey, binding, style = 'marker') => {
        const key = `${propKey}|${binding.key}|${style}`;
        if (scatterChartGroups.has(key)) return scatterChartGroups.get(key);
        const group = propKey === 'c:bubbleChart'
            ? createBubbleChart(binding.xId, binding.yId)
            : createScatterChart(binding.xId, binding.yId, style);
        appendPlotAreaItem(plotArea, propKey, group);
        scatterChartGroups.set(key, group);
        return group;
    };

    seriesList.forEach((s, index) => {
        if (s.type == 'bar') {
            const barDir = (xAxis.type || 'category') === 'category' ? 'col' : 'bar';
            const isStack = s.stack == 'total';
            const binding = getCategoryAxisBinding(s, barDir);
            const valueAxis = binding.valueAxis;
            const isPercent = isStack && isPercentStackedChart(chartOption, valueAxis);
            const group = getChartGroup(
                'c:barChart',
                `${binding.categoryIndex}|${binding.valueIndex}|${barDir}|${isStack ? 1 : 0}|${isPercent ? 1 : 0}`,
                () => createBarChart(binding.categoryId, binding.valueId, isStack, barDir, isPercent)
            );
            group["c:ser"].push({
                "c:idx": { "_$val": index }, "c:order": { "_$val": index },
                "c:tx": buildSeriesText(s.name),
                "c:invertIfNegative": { "_$val": "1" },
                ...buildBarStyle(s),
                ...(buildSeriesDataLabels(s) || {}),
                "c:cat": buildRef(binding.categories, 'str'),
                "c:val": buildRef(s.data, 'num'),
                "c:extLst": buildSeriesExt(index, chartGuid)
            });
            hasChart = true;
        } else if (s.type == 'line') {
            const chartType = s.areaStyle ? 'c:areaChart' : 'c:lineChart';
            const binding = getCategoryAxisBinding(s, (xAxis.type || 'category') === 'category' ? 'col' : 'bar');
            const isPercent = s.stack && isPercentStackedChart(chartOption, binding.valueAxis);
            const grouping = s.stack ? (isPercent ? 'percentStacked' : 'stacked') : 'standard';
            const group = getChartGroup(
                chartType,
                `${binding.categoryIndex}|${binding.valueIndex}|${grouping}`,
                () => createLineChart(binding.categoryId, binding.valueId, grouping)
            );
            group["c:ser"].push({
                "c:idx": { "_$val": index }, "c:order": { "_$val": index },
                "c:tx": buildSeriesText(s.name),
                ...buildLineMarker(s),
                ...buildLineStyle(s),
                ...(buildSeriesDataLabels(s) || {}),
                "c:cat": buildRef(binding.categories, 'str'),
                "c:val": buildRef(s?.data, 'num'),
                "c:smooth": { "_$val": s?.smooth ? "1" : "0" },
                "c:extLst": buildSeriesExt(index, chartGuid)
            });
            hasChart = true;
        } else if (s.type == 'scatter' && s.__bubble) {
            const binding = getScatterAxisBinding(s);
            const bubbleChart = getScatterChartGroup('c:bubbleChart', binding);
            const xVals = (s.data || []).map(item => Array.isArray(item) ? item[0] : item);
            const yVals = (s.data || []).map(item => Array.isArray(item) ? item[1] : item);
            const sizeVals = extractBubbleSizes(s.data || []);
            bubbleChart["c:ser"].push({
                "c:idx": { "_$val": index }, "c:order": { "_$val": index },
                "c:tx": buildSeriesText(s.name),
                ...buildBubbleStyle(s),
                ...(buildSeriesDataLabels(s) || {}),
                "c:xVal": buildRef(xVals, 'num'),
                "c:yVal": buildRef(yVals, 'num'),
                "c:bubbleSize": buildRef(sizeVals, 'num'),
                "c:bubble3D": { "_$val": "0" },
                "c:extLst": buildSeriesExt(index, chartGuid)
            });
            hasChart = true;
        } else if (s.type == 'scatter') {
            const binding = getScatterAxisBinding(s);
            const xVals = (s.data || []).map(item => Array.isArray(item) ? item[0] : item);
            const yVals = (s.data || []).map(item => Array.isArray(item) ? item[1] : item);
            const scatterStyle = getScatterStyle(s);
            const scatterChart = getScatterChartGroup('c:scatterChart', binding, scatterStyle);
            scatterChart["c:ser"].push({
                "c:idx": { "_$val": index }, "c:order": { "_$val": index },
                "c:tx": buildSeriesText(s.name),
                ...buildScatterMarker(s),
                ...buildScatterLineStyle(scatterStyle, s),
                ...buildScatterSmooth(scatterStyle),
                ...(buildSeriesDataLabels(s) || {}),
                "c:xVal": buildRef(xVals, 'num'),
                "c:yVal": buildRef(yVals, 'num'),
                "c:extLst": buildSeriesExt(index, chartGuid)
            });
            hasChart = true;
        } else if (s.type == 'radar') {
            const binding = getCategoryAxisBinding(s);
            const group = getChartGroup(
                'c:radarChart',
                `${binding.categoryIndex}|${binding.valueIndex}|${radarChartStyle}`,
                () => createRadarChart(binding.categoryId, binding.valueId, radarChartStyle)
            );
            const rawRadarCats = radar?.indicator?.map(item => item?.name) || xAxis?.data || [];
            const radarCats = Array.isArray(rawRadarCats) ? processCellReferences(SN, rawRadarCats) : rawRadarCats;
            const radarDataRaw = Array.isArray(s.data)
                ? ((s.data[0] && typeof s.data[0] === 'object' && !Array.isArray(s.data[0]) && Object.prototype.hasOwnProperty.call(s.data[0], 'value'))
                    ? s.data[0].value
                    : s.data)
                : s.data;
            const radarData = Array.isArray(radarDataRaw) ? processCellReferences(SN, radarDataRaw) : radarDataRaw;
            group["c:ser"].push({
                "c:idx": { "_$val": index }, "c:order": { "_$val": index },
                "c:tx": buildSeriesText(s.name),
                ...buildLineMarker(s),
                ...buildLineStyle(s),
                ...(buildSeriesDataLabels(s) || {}),
                "c:cat": buildRef(radarCats, 'str'),
                "c:val": buildRef(radarData, 'num'),
                "c:extLst": buildSeriesExt(index, chartGuid)
            });
            hasChart = true;
        } else if (s.type == 'pie') {
            const chartType = Array.isArray(s.radius) ? 'c:doughnutChart' : 'c:pieChart';
            if (!plotArea['c:pieChart'] && !plotArea['c:doughnutChart']) {
                plotArea[chartType] = createPieChart(s);
            }
            plotArea[chartType]['c:ser'] = {
                "c:idx": { "_$val": index }, "c:order": { "_$val": index },
                "c:tx": buildSeriesText(s.name),
                ...(buildSeriesDataLabels(s) || {}),
                "c:dPt": s.data.map((obj, idx) => ({
                    "c:idx": { "_$val": idx },
                    "c:bubble3D": { "_$val": "0" },
                    ...buildDataPointStyle(obj)
                })),
                "c:cat": buildRef(processCellReferences(SN, s.data.map(obj => obj.name)), 'str'),
                "c:val": buildRef(processCellReferences(SN, s.data.map(obj => obj.value)), 'num'),
                "c:extLst": buildSeriesExt(index, chartGuid)
            };
            hasChart = true;
        }
    });

    if (!hasChart && seriesList.length > 0) {
        const fallback = seriesList[0];
        const binding = getCategoryAxisBinding(fallback);
        const group = getChartGroup(
            'c:barChart',
            `${binding.categoryIndex}|${binding.valueIndex}|col|0|0`,
            () => createBarChart(binding.categoryId, binding.valueId, false, 'col')
        );
        group["c:ser"].push({
            "c:idx": { "_$val": 0 }, "c:order": { "_$val": 0 },
            "c:tx": buildSeriesText(fallback.name),
            "c:invertIfNegative": { "_$val": "1" },
            ...(buildSeriesDataLabels(fallback) || {}),
            "c:cat": buildRef(binding.categories, 'str'),
            "c:val": buildRef(fallback.data ?? [], 'num'),
            "c:extLst": buildSeriesExt(0, chartGuid)
        });
    }

    flushCategoryAxes();
    const plotAreaSpPr = buildPlotAreaSpPr(chartOption || {});
    if (plotAreaSpPr) {
        plotArea["c:spPr"] = plotAreaSpPr;
    }

    const chartSpace = createChartSpace({ ...d, chartOption }, plotArea);
    normalizeChartSpaceNode(chartSpace);

    return {
        "?xml": { "_$version": "1.0", "_$encoding": "UTF-8", "_$standalone": "yes" },
        "c:chartSpace": chartSpace
    };
}

export { createGuid };
