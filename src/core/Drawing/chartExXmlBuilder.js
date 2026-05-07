/**
 * ChartEx XML Export Module
 * Provides build capabilities for Office 2014 + extended diagrams (sunrise diagrams, tree diagrams, waterfall diagrams, histograms, etc.)
 */

import { CHART_EX_STYLE_OBJ_MAP, CHART_EX_COLOR_OBJ } from "../IO/chartExTemplates.js";
import { createGuid } from "./chartXmlBuilder.js";

/**
 * Clone ChartEx Style
 */
export function cloneChartExStyle(type) {
    const base = CHART_EX_STYLE_OBJ_MAP[type] || CHART_EX_STYLE_OBJ_MAP.sunburst;
    return JSON.parse(JSON.stringify(base));
}

/**
 * Clone ChartEx Color
 */
export function cloneChartExColors() {
    return JSON.parse(JSON.stringify(CHART_EX_COLOR_OBJ));
}

/**
 * Create ChartEx Context
 */
export function createChartExContext(_Xml) {
    const workbookRoot = _Xml.obj["xl/workbook.xml"];
    let workbook = workbookRoot.workbook;
    if (!workbook.definedNames) {
        workbook.definedNames = { definedName: [] };
    }
    const list = Array.isArray(workbook.definedNames.definedName)
        ? workbook.definedNames.definedName
        : (workbook.definedNames.definedName ? [workbook.definedNames.definedName] : []);
    workbook.definedNames.definedName = list;
    const keys = Object.keys(workbook);
    const calcIndex = keys.indexOf("calcPr");
    const definedIndex = keys.indexOf("definedNames");
    if (calcIndex !== -1 && definedIndex !== -1 && definedIndex > calcIndex) {
        const reordered = {};
        for (const key of keys) {
            if (key === "calcPr") {
                reordered.definedNames = workbook.definedNames;
                reordered.calcPr = workbook.calcPr;
                continue;
            }
            if (key === "definedNames") continue;
            reordered[key] = workbook[key];
        }
        if (!reordered.definedNames) reordered.definedNames = workbook.definedNames;
        workbookRoot.workbook = reordered;
        workbook = reordered;
    }

    let maxVersion = 0;
    list.forEach((item) => {
        const name = item?._$name;
        const match = typeof name === 'string' ? name.match(/^_xlchart\.v(\d+)\./) : null;
        if (match) {
            const ver = Number(match[1]);
            if (Number.isFinite(ver)) maxVersion = Math.max(maxVersion, ver);
        }
    });
    let nextVersion = maxVersion + 1;

    return {
        nextPrefix() {
            return `_xlchart.v${nextVersion++}`;
        },
        addDefinedName(name, ref) {
            list.push({
                "_$name": name,
                "_$hidden": "1",
                "#text": ref
            });
        }
    };
}

/**
 * Parse ChartEx Export Information
 */
export function resolveChartExExportInfo(d) {
    const chartOption = d.chartOption || {};
    const series = Array.isArray(chartOption.series) ? chartOption.series : [];
    const typeFromOption = chartOption.__chartType || (chartOption.__histogram ? 'histogram' : null);
    const typeFromSeries = series.find(s => ['sunburst', 'treemap', 'funnel'].includes(s?.type))?.type;
    const chartType = typeFromOption || typeFromSeries;
    if (!['sunburst', 'treemap', 'funnel', 'waterfall', 'histogram'].includes(chartType)) return null;

    const excelRef = chartOption.__excelRef;
    if (!excelRef?.categories || !Array.isArray(excelRef.series) || excelRef.series.length === 0) return null;

    return {
        type: chartType,
        categories: excelRef.categories,
        series: excelRef.series,
        chartOption,
        hiddenFlags: chartOption.__excelHidden,
        histogram: chartOption.__histogram
    };
}

/**
 * Building the ChartEx Graphics Framework
 */
export function buildChartExGraphicFrame(i, chartName) {
    return {
        "xdr:nvGraphicFramePr": { "xdr:cNvPr": { "_$id": i + 1, "_$name": `图表 ${i + 1}` }, "xdr:cNvGraphicFramePr": "" },
        "xdr:xfrm": { "a:off": { "_$x": "0", "_$y": "0" }, "a:ext": { "_$cx": "0", "_$cy": "0" } },
        "a:graphic": {
            "a:graphicData": {
                "_$uri": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                "cx:chart": {
                    "_$xmlns:cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
                    "_$xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "_$r:id": 'rId' + (i + 1)
                }
            }
        },
        "_$macro": ""
    };
}

/**
 * Building the ChartEx Relationship File
 */
export function buildChartExRels(index) {
    return {
        "?xml": { "_$version": "1.0", "_$encoding": "UTF-8", "_$standalone": "yes" },
        "Relationships": {
            "Relationship": [
                { "_$Id": "rId1", "_$Type": "http://schemas.microsoft.com/office/2011/relationships/chartStyle", "_$Target": `style${index}.xml` },
                { "_$Id": "rId2", "_$Type": "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle", "_$Target": `colors${index}.xml` }
            ],
            "_$xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
        }
    };
}

/**
 * Get ChartEx text value
 */
function getChartExTextValue(SN, refOrValue) {
    if (typeof refOrValue !== 'string') return '';
    if (!refOrValue.includes('!')) return refOrValue;
    const [sheetNameRaw, cellRefRaw] = refOrValue.split('!');
    const sheetName = sheetNameRaw.replaceAll('\'', '');
    const sheet = SN.getSheet(sheetName);
    if (!sheet) return '';
    const cellRef = cellRefRaw.replaceAll('$', '');
    if (cellRef.includes(':')) {
        const first = cellRef.split(':')[0];
        return sheet.getCell(first).calcVal;
    }
    return sheet.getCell(cellRef).calcVal;
}

/**
 * Build ChartEx Header
 */
function buildChartExTitle(titleText) {
    return {
        "_$pos": "t",
        "_$align": "ctr",
        "_$overlay": "0",
        "cx:tx": { "cx:txData": { "cx:v": titleText } },
        "cx:txPr": {
            "a:bodyPr": { "_$spcFirstLastPara": "1", "_$vertOverflow": "ellipsis", "_$horzOverflow": "overflow", "_$wrap": "square", "_$lIns": "0", "_$tIns": "0", "_$rIns": "0", "_$bIns": "0", "_$anchor": "ctr", "_$anchorCtr": "1" },
            "a:lstStyle": "",
            "a:p": {
                "a:pPr": { "_$algn": "ctr", "_$rtl": "0", "a:defRPr": "" },
                "a:r": [{
                    "a:rPr": { "_$lang": "zh-CN", "_$altLang": "en-US", "_$sz": "1400", "_$b": "0", "_$i": "0", "_$u": "none", "_$strike": "noStrike", "_$baseline": "0", "a:solidFill": { "a:sysClr": { "_$val": "windowText", "_$lastClr": "000000", "a:lumMod": { "_$val": "65000" }, "a:lumOff": { "_$val": "35000" } } }, "a:latin": { "_$typeface": "Calibri" }, "a:ea": { "_$typeface": "宋体", "_$panose": "02010600030101010101", "_$pitchFamily": "2", "_$charset": "-122" } },
                    "a:t": titleText
                }]
            }
        }
    };
}

/**
 * Build ChartEx Legend
 */
function buildChartExLegend(legend) {
    if (legend?.show === false) return null;
    let pos = 'b';
    if (legend?.top != null) pos = 't';
    if (legend?.left != null) pos = 'l';
    if (legend?.right != null) pos = 'r';
    return { "_$pos": pos, "_$align": "ctr", "_$overlay": "0" };
}

/**
 * Building the ChartEx Series
 */
function buildChartExSeries(info, seriesRefs, defineRef, SN) {
    const layoutId = info.type === 'histogram' ? 'clusteredColumn' : info.type;
    const isHistogram = info.type === 'histogram';
    const useSize = layoutId === 'sunburst' || layoutId === 'treemap';
    const hiddenFlags = Array.isArray(info.hiddenFlags) ? info.hiddenFlags : null;
    let dataLabel = null;

    if (layoutId === 'treemap') {
        dataLabel = { "_$pos": "inEnd", "cx:visibility": { "_$seriesName": "0", "_$categoryName": "1", "_$value": "0" } };
    } else if (layoutId === 'sunburst') {
        dataLabel = { "_$pos": "ctr", "cx:visibility": { "_$seriesName": "0", "_$categoryName": "1", "_$value": "0" } };
    } else if (layoutId === 'waterfall') {
        dataLabel = { "_$pos": "outEnd", "cx:visibility": { "_$seriesName": "0", "_$categoryName": "0", "_$value": "1" } };
    } else if (!isHistogram) {
        dataLabel = { "cx:visibility": { "_$seriesName": "0", "_$categoryName": "0", "_$value": "1" } };
    }

    return seriesRefs.map((ser, index) => {
        const nameRef = defineRef(ser.name);
        const dataRef = defineRef(ser.data);
        const series = {
            "_$layoutId": layoutId,
            "_$uniqueId": `{${createGuid()}}`,
            "_$formatIdx": String(index),
            "cx:tx": { "cx:txData": { "cx:f": nameRef, "cx:v": String(getChartExTextValue(SN, ser.name) ?? '') } }
        };
        if (dataLabel) series["cx:dataLabels"] = dataLabel;
        series["cx:dataId"] = { "_$val": String(index) };
        if (layoutId === 'treemap') {
            series["cx:layoutPr"] = { "cx:parentLabelLayout": { "_$val": "overlapping" } };
        }
        if (layoutId === 'waterfall') {
            series["cx:layoutPr"] = { "cx:subtotals": "" };
        }
        if (isHistogram) {
            const intervalClosed = info.histogram?.intervalClosed || 'r';
            series["cx:layoutPr"] = { "cx:binning": { "_$intervalClosed": intervalClosed } };
        }
        if (hiddenFlags ? hiddenFlags[index] : index > 0) series["_$hidden"] = "1";
        return { series, dataRef, useSize };
    });
}

/**
 * Get ChartEx XML Details
 * @param {Workbook} SN - Workbook Instances
 * @param {Drawing} d - Chart Drawing Objects
 * @param {Object} info - ChartEx Export Information
 * @param {Object} context - ChartEx Context
 * @ returns {Object} ChartEx XML object
 */
export function getChartExXmlDetail(SN, d, info, context) {
    const option = info.chartOption || {};
    const prefix = context.nextPrefix();
    const refCache = new Map();
    let refIndex = 0;
    const defineRef = (ref) => {
        if (refCache.has(ref)) return refCache.get(ref);
        const name = `${prefix}.${refIndex++}`;
        context.addDefinedName(name, ref);
        refCache.set(ref, name);
        return name;
    };

    const seriesPack = buildChartExSeries(info, info.series, defineRef, SN);
    const catRef = defineRef(info.categories);
    const chartData = {
        "cx:data": seriesPack.map((pack, index) => ({
            "_$id": String(index),
            "cx:strDim": { "_$type": "cat", "cx:f": { "_$dir": "row", "#text": catRef } },
            "cx:numDim": { "_$type": pack.useSize ? "size" : "val", "cx:f": { "_$dir": "row", "#text": pack.dataRef } }
        }))
    };

    const plotArea = {
        "cx:plotAreaRegion": { "cx:series": seriesPack.map(pack => pack.series) }
    };

    if (info.type === 'funnel') {
        plotArea["cx:axis"] = { "_$id": "0", "cx:catScaling": { "_$gapWidth": "0.0599999987" }, "cx:tickLabels": "" };
    } else if (info.type === 'waterfall') {
        plotArea["cx:axis"] = [
            { "_$id": "0", "cx:catScaling": { "_$gapWidth": "0.5" }, "cx:tickLabels": "" },
            { "_$id": "1", "cx:valScaling": "", "cx:majorGridlines": "", "cx:tickLabels": "" }
        ];
    } else if (info.type === 'histogram') {
        plotArea["cx:axis"] = [
            { "_$id": "0", "cx:catScaling": { "_$gapWidth": "0" }, "cx:tickLabels": "" },
            { "_$id": "1", "cx:valScaling": "", "cx:majorGridlines": "", "cx:tickLabels": "" }
        ];
    }

    const titleText = option?.title?.text ?? '';
    const legendObj = buildChartExLegend(option.legend);
    const chart = {
        "cx:title": buildChartExTitle(titleText),
        "cx:plotArea": plotArea
    };
    if (legendObj) chart["cx:legend"] = legendObj;

    return {
        "?xml": { "_$version": "1.0", "_$encoding": "UTF-8", "_$standalone": "yes" },
        "cx:chartSpace": {
            "_$xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "_$xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "_$xmlns:cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
            "cx:chartData": chartData,
            "cx:chart": chart
        }
    };
}
