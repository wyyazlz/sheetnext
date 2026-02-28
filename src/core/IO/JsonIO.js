
import { IndexTable, isCellDataEqual } from "./helpers.js";
import AutoFilter from "../AutoFilter/AutoFilter.js";
import Table from "../Table/Table.js";
import CF from "../CF/CF.js";
import Drawing from "../Drawing/Drawing.js";
import Comment from "../Comment/Comment.js";
import Sparkline from "../Sparkline/Sparkline.js";
import { PivotTable, PivotTableItem, PivotCache } from "../PivotTable/index.js";
import { Slicer } from "../Slicer/index.js";
import SlicerItem from "../Slicer/SlicerItem.js";
import SlicerCache from "../Slicer/SlicerCache.js";
import SheetProtection from "../Sheet/SheetProtection.js";

const JSON_VERSION = "2.0";
const DEFAULT_ROW_COUNT = 200;
const DEFAULT_COL_COUNT = 32;
const DYNAMIC_TYPE_KEY = "__sn_json_type";
const DYNAMIC_VALUE_KEY = "__sn_json_value";

function _isObject(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
}

function _clone(value) {
    if (value === undefined) return undefined;
    return JSON.parse(JSON.stringify(value));
}

function _safeInt(value, fallback = 0) {
    const num = Number.parseInt(value, 10);
    return Number.isFinite(num) ? num : fallback;
}

function _safeNumber(value, fallback = null) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
}

function _normalizeCellRef(ref, fallback = { r: 0, c: 0 }) {
    if (!_isObject(ref)) return { r: fallback.r, c: fallback.c };
    const r = Math.max(0, _safeInt(ref.r, fallback.r));
    const c = Math.max(0, _safeInt(ref.c, fallback.c));
    return { r, c };
}

function _normalizeRange(range) {
    if (!range?.s || !range?.e) return null;
    const sr = Math.max(0, _safeInt(range.s.r, 0));
    const sc = Math.max(0, _safeInt(range.s.c, 0));
    const er = Math.max(0, _safeInt(range.e.r, sr));
    const ec = Math.max(0, _safeInt(range.e.c, sc));
    return {
        s: { r: Math.min(sr, er), c: Math.min(sc, ec) },
        e: { r: Math.max(sr, er), c: Math.max(sc, ec) }
    };
}

function _normalizeZoom(zoom) {
    const num = Number(zoom);
    if (!Number.isFinite(num)) return 1;
    return Math.min(4, Math.max(0.1, num));
}

function _normalizeFormulaForImport(formula) {
    if (typeof formula !== 'string') return null;
    const text = formula.trim();
    if (!text) return null;
    return text.startsWith('=') ? text : `=${text}`;
}

function _encodeCellValue(value, cellType = null) {
    if (value instanceof Date) {
        return {
            value: value.toISOString(),
            type: ['date', 'time', 'dateTime'].includes(cellType) ? cellType : 'dateTime'
        };
    }

    if (typeof value === 'number') {
        if (Number.isNaN(value)) return { value: null, type: 'nan' };
        if (!Number.isFinite(value)) return { value: null, type: value > 0 ? 'inf' : '-inf' };
    }

    return { value, type: null };
}

function _decodeCellValue(value, type) {
    if (!type) return value;

    if (type === 'nan') return Number.NaN;
    if (type === 'inf') return Number.POSITIVE_INFINITY;
    if (type === '-inf') return Number.NEGATIVE_INFINITY;
    if (type === 'date' || type === 'time' || type === 'dateTime') {
        if (!value) return null;
        const date = new Date(value);
        return Number.isNaN(date.getTime()) ? null : date;
    }

    return value;
}

function _cloneCellPrimitive(value) {
    if (value instanceof Date) return new Date(value.getTime());
    if (Array.isArray(value) || _isObject(value)) return _clone(value);
    return value;
}

function _serializeDynamicValue(value) {
    if (value instanceof Date) {
        return {
            [DYNAMIC_TYPE_KEY]: 'date',
            [DYNAMIC_VALUE_KEY]: value.toISOString()
        };
    }

    if (typeof value === 'number') {
        if (Number.isNaN(value)) {
            return {
                [DYNAMIC_TYPE_KEY]: 'nan',
                [DYNAMIC_VALUE_KEY]: null
            };
        }
        if (!Number.isFinite(value)) {
            return {
                [DYNAMIC_TYPE_KEY]: value > 0 ? 'inf' : '-inf',
                [DYNAMIC_VALUE_KEY]: null
            };
        }
        return value;
    }

    if (Array.isArray(value)) {
        return value.map(item => _serializeDynamicValue(item));
    }

    if (_isObject(value)) {
        const output = {};
        Object.keys(value).forEach(key => {
            output[key] = _serializeDynamicValue(value[key]);
        });
        return output;
    }

    return value;
}

function _deserializeDynamicValue(value) {
    if (Array.isArray(value)) {
        return value.map(item => _deserializeDynamicValue(item));
    }

    if (_isObject(value) && Object.prototype.hasOwnProperty.call(value, DYNAMIC_TYPE_KEY)) {
        const type = value[DYNAMIC_TYPE_KEY];
        const raw = value[DYNAMIC_VALUE_KEY];
        if (type === 'date') {
            const date = new Date(raw);
            return Number.isNaN(date.getTime()) ? null : date;
        }
        if (type === 'nan') return Number.NaN;
        if (type === 'inf') return Number.POSITIVE_INFINITY;
        if (type === '-inf') return Number.NEGATIVE_INFINITY;
        return raw;
    }

    if (_isObject(value)) {
        const output = {};
        Object.keys(value).forEach(key => {
            output[key] = _deserializeDynamicValue(value[key]);
        });
        return output;
    }

    return value;
}
function _resolveSheetSize(sheetData) {
    let rowCount = Math.max(_safeInt(sheetData?.rowCount, DEFAULT_ROW_COUNT), 1);
    let colCount = Math.max(_safeInt(sheetData?.colCount, DEFAULT_COL_COUNT), 1);

    const rows = Array.isArray(sheetData?.rows) ? sheetData.rows : [];
    rows.forEach(rowData => {
        const rIndex = Math.max(0, _safeInt(rowData?.rIndex, 0));
        rowCount = Math.max(rowCount, rIndex + 1);
        if (!Array.isArray(rowData?.cells)) return;
        rowData.cells.forEach(cellData => {
            const startCol = Math.max(0, _safeInt(cellData?.c, 0));
            const endCol = Math.max(startCol, _safeInt(cellData?.e, startCol));
            colCount = Math.max(colCount, endCol + 1);
        });
    });

    const cols = Array.isArray(sheetData?.cols) ? sheetData.cols : [];
    cols.forEach(colData => {
        const cIndex = Math.max(0, _safeInt(colData?.cIndex, 0));
        colCount = Math.max(colCount, cIndex + 1);
    });

    const merges = Array.isArray(sheetData?.merges) ? sheetData.merges : [];
    merges.forEach(merge => {
        const range = _normalizeRange(merge);
        if (!range) return;
        rowCount = Math.max(rowCount, range.e.r + 1);
        colCount = Math.max(colCount, range.e.c + 1);
    });

    const activeCell = _normalizeCellRef(sheetData?.activeCell, { r: 0, c: 0 });
    const viewStart = _normalizeCellRef(sheetData?.viewStart, { r: 0, c: 0 });

    rowCount = Math.max(rowCount, activeCell.r + 1, viewStart.r + 1);
    colCount = Math.max(colCount, activeCell.c + 1, viewStart.c + 1);

    rowCount = Math.max(rowCount, _safeInt(sheetData?.frozenRows, 0));
    colCount = Math.max(colCount, _safeInt(sheetData?.frozenCols, 0));
    rowCount = Math.max(rowCount, _safeInt(sheetData?.freezeStartRow, 0) + 1);
    colCount = Math.max(colCount, _safeInt(sheetData?.freezeStartCol, 0) + 1);

    return { rowCount, colCount };
}

function _ensureSheetSize(sheet, rowCount, colCount) {
    const safeRows = Math.max(1, rowCount);
    const safeCols = Math.max(1, colCount);
    sheet.getRow(safeRows - 1);
    sheet.getCol(safeCols - 1);
}

function _resetSheetForImport(sheet) {
    sheet.rows = [];
    sheet.cols = [];
    sheet.merges = [];
    sheet.vi = { rowArr: [0], colArr: [0] };
    sheet.views = [{ pane: {} }];
    sheet._activeAreas = [];
    sheet._brush = { keep: false, area: null };
    sheet._styledRows = new Set();
    sheet._styledCols = new Set();
    sheet._rowVisibilityVersion = 0;
    sheet._totalWidthCache = undefined;
    sheet._totalHeightCache = undefined;
    sheet._idCount = 0;

    sheet.protection = new SheetProtection(sheet);
    sheet.Comment = new Comment(sheet);
    sheet.AutoFilter = new AutoFilter(sheet);
    sheet.Table = new Table(sheet);
    sheet.Drawing = new Drawing(sheet);
    sheet.Sparkline = new Sparkline(sheet);
    sheet.CF = new CF(sheet);
    sheet.Slicer = new Slicer(sheet);
    sheet.PivotTable = new PivotTable(sheet);
}

function _buildCellStyleFromData(cellData, styleTable) {
    let style = null;

    if (cellData?.s !== undefined && styleTable?.styles?.[cellData.s]) {
        style = _clone(styleTable.styles[cellData.s]);
    } else {
        style = {};
    }

    if (cellData?.fo !== undefined && styleTable?.fonts?.[cellData.fo]) {
        style.font = _clone(styleTable.fonts[cellData.fo]);
    }
    if (cellData?.fi !== undefined && styleTable?.fills?.[cellData.fi]) {
        style.fill = _clone(styleTable.fills[cellData.fi]);
    }
    if (cellData?.b !== undefined && styleTable?.borders?.[cellData.b]) {
        style.border = _clone(styleTable.borders[cellData.b]);
    }
    if (cellData?.a !== undefined && styleTable?.aligns?.[cellData.a]) {
        style.alignment = _clone(styleTable.aligns[cellData.a]);
    }
    if (cellData?.fmt !== undefined && cellData?.fmt !== null) {
        style.numFmt = cellData.fmt;
    }
    if (cellData?.p) {
        style.protection = _clone(cellData.p);
    }

    if (!style || Object.keys(style).length === 0) return null;
    return style;
}

function _createDecodedCellTemplate(cellData, styleTable) {
    const formula = _normalizeFormulaForImport(cellData?.f);
    const template = {
        hasFormula: !!formula,
        formula,
        hasValue: Object.prototype.hasOwnProperty.call(cellData || {}, 'v') || Object.prototype.hasOwnProperty.call(cellData || {}, 'vt'),
        value: _decodeCellValue(cellData?.v, cellData?.vt),
        valueType: cellData?.vt ?? null,
        calcValue: _decodeCellValue(cellData?.cv, cellData?.cvt),
        calcType: cellData?.cvt ?? null,
        style: _buildCellStyleFromData(cellData, styleTable),
        link: cellData?.link ? {
            target: cellData.link.t ?? null,
            location: cellData.link.l ?? null,
            tooltip: cellData.link.tt ?? null
        } : null,
        dataValidation: cellData?.dv ? _clone(cellData.dv) : null,
        richText: Array.isArray(cellData?.rt) ? _clone(cellData.rt) : null
    };

    return template;
}

function _applyDecodedCellTemplate(cell, template) {
    cell.isMerged = false;
    cell.master = null;
    cell._showVal = undefined;
    cell._fmtResult = undefined;
    cell._spillRange = null;
    cell._spillParent = null;
    cell._spillOffset = null;
    cell._isVolatile = false;

    if (template.hasFormula) {
        cell._editVal = template.formula;
        cell._calcVal = template.calcValue !== undefined ? _cloneCellPrimitive(template.calcValue) : undefined;
        if (template.calcType && ['date', 'time', 'dateTime'].includes(template.calcType)) {
            cell._type = template.calcType;
        } else {
            cell._type = undefined;
        }
    } else if (template.hasValue) {
        cell._editVal = _cloneCellPrimitive(template.value);
        cell._calcVal = undefined;
        if (template.valueType && ['date', 'time', 'dateTime'].includes(template.valueType)) {
            cell._type = template.valueType;
        } else if (typeof template.value === 'boolean') {
            cell._type = 'boolean';
        } else if (typeof template.value === 'string') {
            cell._type = 'string';
        } else {
            cell._type = undefined;
        }
    } else {
        cell._editVal = "";
        cell._calcVal = undefined;
        cell._type = undefined;
    }

    if (template.richText) {
        cell._richText = _clone(template.richText);
        if ((cell._editVal === null || cell._editVal === undefined || cell._editVal === "") && Array.isArray(template.richText)) {
            cell._editVal = template.richText.map(run => run?.text ?? '').join('');
            if (typeof cell._editVal === 'string') cell._type = 'string';
        }
    } else {
        cell._richText = null;
    }

    cell._style = template.style ? _clone(template.style) : {};
    cell._hyperlink = template.link ? { ...template.link } : null;
    cell._dataValidation = template.dataValidation ? _clone(template.dataValidation) : null;
}

function _serializeCfRules(sheet) {
    if (!sheet.CF?.getAll) return [];
    return sheet.CF.getAll().map(rule => ({
        ..._clone(rule._config || {}),
        type: rule.type,
        priority: rule.priority,
        sqref: rule.sqref,
        stopIfTrue: !!rule.stopIfTrue,
        dxf: rule.dxf ? _clone(rule.dxf) : null
    }));
}

function _restoreCfRules(sheet, cfRules) {
    if (!sheet.CF) return;
    sheet.CF.rules = [];
    sheet.CF.clearAllCache();

    if (!Array.isArray(cfRules)) return;
    cfRules.forEach(ruleData => {
        if (!ruleData) return;
        const config = _clone(ruleData);
        const rangeRef = config.sqref;
        if (!rangeRef) return;
        sheet.CF.add({
            ...config,
            rangeRef
        });
    });
}
function _serializePivotCache(cache) {
    return {
        cacheId: cache.cacheId,
        sourceType: cache.sourceType,
        sourceRef: cache.sourceRef,
        sourceSheet: cache.sourceSheet,
        sourceRange: _clone(cache.sourceRange),
        fields: _serializeDynamicValue(cache.fields || []),
        records: _serializeDynamicValue(cache.records || []),
        recordCount: Number(cache.recordCount || 0),
        refreshedBy: cache.refreshedBy || '',
        refreshedDate: cache.refreshedDate ?? null,
        createdVersion: Number(cache.createdVersion || 7),
        refreshedVersion: Number(cache.refreshedVersion || 8),
        refreshOnLoad: !!cache.refreshOnLoad
    };
}

function _restorePivotCaches(SN, pivotCachesData) {
    SN._pivotCaches = new Map();
    if (!Array.isArray(pivotCachesData)) return;

    pivotCachesData.forEach(cacheData => {
        if (!cacheData) return;
        const cacheId = _safeInt(cacheData.cacheId, null);
        const cache = new PivotCache(SN, {
            cacheId: cacheId ?? cacheData.cacheId,
            sourceType: cacheData.sourceType || 'worksheet',
            sourceRef: cacheData.sourceRef ?? null,
            sourceSheet: cacheData.sourceSheet ?? null,
            sourceRange: _clone(cacheData.sourceRange),
            refreshedBy: cacheData.refreshedBy || '',
            refreshedDate: cacheData.refreshedDate ?? null,
            createdVersion: _safeInt(cacheData.createdVersion, 7),
            refreshedVersion: _safeInt(cacheData.refreshedVersion, 8),
            refreshOnLoad: !!cacheData.refreshOnLoad
        });

        cache.fields = _deserializeDynamicValue(cacheData.fields || []);
        cache.records = _deserializeDynamicValue(cacheData.records || []);
        cache.recordCount = _safeInt(cacheData.recordCount, cache.records.length);
        if (typeof cache._hydrateSharedItemsFromRecords === 'function') {
            cache._hydrateSharedItemsFromRecords();
        }
        SN._pivotCaches.set(cache.cacheId, cache);
    });
}

function _serializeSlicerCache(cache) {
    const items = Array.isArray(cache.items) ? cache.items : [];
    return {
        cacheId: cache.cacheId,
        sourceType: cache.sourceType,
        sourceSheet: cache.sourceSheet,
        sourceName: cache.sourceName,
        sourceId: cache.sourceId ?? null,
        fieldIndex: cache.fieldIndex,
        fieldName: cache.fieldName,
        items: _serializeDynamicValue(items)
    };
}

function _restoreSlicerCaches(SN, slicerCachesData) {
    SN._slicerCaches = new Map();
    if (!Array.isArray(slicerCachesData)) return;

    slicerCachesData.forEach(cacheData => {
        if (!cacheData) return;
        const cacheId = _safeInt(cacheData.cacheId, null);
        const cache = new SlicerCache(SN, {
            cacheId: cacheId ?? cacheData.cacheId,
            sourceType: cacheData.sourceType || 'table',
            sourceSheet: cacheData.sourceSheet ?? null,
            sourceName: cacheData.sourceName ?? null,
            sourceId: cacheData.sourceId ?? null,
            fieldIndex: _safeInt(cacheData.fieldIndex, 0),
            fieldName: cacheData.fieldName || ''
        });

        const items = _deserializeDynamicValue(cacheData.items || []);
        if (Array.isArray(items) && items.length > 0) {
            cache.items = items;
            cache._itemMap = new Map(items.map(item => [item.key, item]));
            cache._dirty = false;
            cache._version = 1;
        } else {
            cache.items = [];
            cache._itemMap = new Map();
            cache._dirty = true;
        }

        SN._slicerCaches.set(cache.cacheId, cache);
    });
}

function _serializeDrawing(drawing) {
    const drawingData = {
        type: drawing.type,
        startCell: _clone(drawing.startCell),
        offsetX: drawing.offsetX,
        offsetY: drawing.offsetY,
        width: drawing.width,
        height: drawing.height,
        anchorType: drawing.anchorType,
        rotation: drawing.rotation
    };

    if (drawing.anchorType === 'twoCell' && drawing._endCell) {
        drawingData.endCell = _clone(drawing._endCell);
        drawingData.endOffsetX = drawing._endOffsetX ?? 0;
        drawingData.endOffsetY = drawing._endOffsetY ?? 0;
    }

    if (drawing.type === 'chart') {
        drawingData.chartOption = _clone(drawing.chartOption);
    } else if (drawing.type === 'image') {
        drawingData.imageBase64 = drawing.imageBase64;
    } else if (drawing.type === 'shape' || drawing.type === 'connector') {
        drawingData.shapeType = drawing.shapeType;
        drawingData.shapeStyle = _clone(drawing.shapeStyle);
        drawingData.shapeText = drawing.shapeText;
        drawingData.isTextBox = !!drawing.isTextBox;
    }

    return drawingData;
}

function _applyDrawingData(sheet, drawingData) {
    if (!drawingData?.type) return;
    if (drawingData.type === 'slicer') return;
    const drawingType = drawingData.type === 'connector' ? 'shape' : drawingData.type;

    const config = {
        type: drawingType,
        startCell: _normalizeCellRef(drawingData.startCell, { r: 0, c: 0 }),
        offsetX: _safeNumber(drawingData.offsetX, 0) ?? 0,
        offsetY: _safeNumber(drawingData.offsetY, 0) ?? 0,
        width: _safeNumber(drawingData.width, null),
        height: _safeNumber(drawingData.height, null),
        anchorType: drawingData.anchorType || 'twoCell',
        rotation: _safeNumber(drawingData.rotation, 0) ?? 0
    };

    if (drawingType === 'chart') {
        config.chartOption = _clone(drawingData.chartOption);
    } else if (drawingType === 'image') {
        config.imageBase64 = drawingData.imageBase64;
    } else if (drawingType === 'shape') {
        config.shapeType = drawingData.shapeType;
        config.shapeStyle = _clone(drawingData.shapeStyle);
        config.shapeText = drawingData.shapeText;
        config.isTextBox = !!drawingData.isTextBox;
    }

    const drawing = sheet.Drawing._add(config);
    if (!drawing) return;

    if (drawing.anchorType === 'twoCell' && drawingData.endCell) {
        drawing._endCell = _normalizeCellRef(drawingData.endCell, drawing.startCell);
        drawing._endOffsetX = _safeNumber(drawingData.endOffsetX, 0) ?? 0;
        drawing._endOffsetY = _safeNumber(drawingData.endOffsetY, 0) ?? 0;
        drawing._area = null;
    }
}

function _serializeSlicer(slicer) {
    const drawing = slicer.Drawing;
    return {
        ...slicer.toJSON(),
        sourceId: slicer.sourceId ?? null,
        drawing: drawing ? {
            startCell: _clone(drawing.startCell),
            offsetX: drawing.offsetX,
            offsetY: drawing.offsetY,
            width: drawing.width,
            height: drawing.height,
            anchorType: drawing.anchorType,
            rotation: drawing.rotation,
            endCell: drawing._endCell ? _clone(drawing._endCell) : null,
            endOffsetX: drawing._endOffsetX ?? 0,
            endOffsetY: drawing._endOffsetY ?? 0
        } : null
    };
}

function _restoreSlicers(SN, pendingSlicers) {
    pendingSlicers.forEach(({ sheet, slicers }) => {
        if (!Array.isArray(slicers) || slicers.length === 0) return;
        slicers.forEach(slicerData => {
            if (!slicerData) return;
            const cacheIdNum = _safeInt(slicerData.cacheId, null);
            const cacheId = cacheIdNum ?? slicerData.cacheId ?? null;

            if (cacheId !== null && !SN._slicerCaches.has(cacheId)) {
                const fallbackCache = new SlicerCache(SN, {
                    cacheId,
                    sourceType: slicerData.sourceType || 'table',
                    sourceSheet: slicerData.sourceSheet || sheet.name,
                    sourceName: slicerData.sourceName ?? null,
                    sourceId: slicerData.sourceId ?? null,
                    fieldIndex: _safeInt(slicerData.fieldIndex, 0),
                    fieldName: slicerData.fieldName || ''
                });
                SN._slicerCaches.set(fallbackCache.cacheId, fallbackCache);
            }

            const drawingData = slicerData.drawing || {};
            const drawing = sheet.Drawing._add({
                type: 'slicer',
                startCell: _normalizeCellRef(
                    drawingData.startCell,
                    { r: 0, c: 0 }
                ),
                offsetX: _safeNumber(drawingData.offsetX, 0) ?? 0,
                offsetY: _safeNumber(drawingData.offsetY, 0) ?? 0,
                width: _safeNumber(drawingData.width, 180) ?? 180,
                height: _safeNumber(drawingData.height, 220) ?? 220,
                anchorType: drawingData.anchorType || 'oneCell',
                rotation: _safeNumber(drawingData.rotation, 0) ?? 0
            });

            if (drawing && drawing.anchorType === 'twoCell' && drawingData.endCell) {
                drawing._endCell = _normalizeCellRef(drawingData.endCell, drawing.startCell);
                drawing._endOffsetX = _safeNumber(drawingData.endOffsetX, 0) ?? 0;
                drawing._endOffsetY = _safeNumber(drawingData.endOffsetY, 0) ?? 0;
                drawing._area = null;
            }

            const slicer = new SlicerItem(sheet, {
                ..._clone(slicerData),
                cacheId,
                selectedKeys: Array.isArray(slicerData.selectedKeys) ? slicerData.selectedKeys : [],
                drawing
            });
            if (slicerData.sourceId !== undefined) {
                slicer.sourceId = slicerData.sourceId;
            }
            sheet.Slicer._addParsed(slicer);

            if (slicer.selectedKeys?.size > 0) {
                slicer.applyFilter();
            }
        });
    });
}
function _restorePivotTables(SN, pendingPivotTables) {
    pendingPivotTables.forEach(({ sheet, pivotTables }) => {
        if (!Array.isArray(pivotTables) || !sheet.PivotTable) return;
        pivotTables.forEach(item => {
            const xml = item?.xml;
            if (!xml) return;
            const cacheId = _safeInt(item.cacheId, null);
            const cache = SN._pivotCaches?.get(cacheId ?? item.cacheId) || SN._pivotCaches?.get(item.cacheId);
            if (!cache) return;
            const pivotTable = PivotTableItem.fromXml(sheet, xml, cache);
            pivotTable._cache = cache;
            sheet.PivotTable._addParsed(pivotTable);
        });
    });
}

function _restoreAutoFilters(pendingAutoFilters) {
    pendingAutoFilters.forEach(({ sheet, scopes }) => {
        if (!Array.isArray(scopes) || scopes.length === 0) return;
        scopes.forEach(scopeData => {
            const xml = scopeData?.xml;
            if (!xml) return;
            const scopeId = scopeData.scopeId ?? null;
            const fallbackRange = _normalizeRange(scopeData.range);
            sheet.AutoFilter.parseScope(scopeId, xml, {
                range: fallbackRange,
                ownerType: scopeData.ownerType,
                ownerId: scopeData.ownerId,
                restoreHiddenRowsFromSheet: false
            });

            if (sheet.AutoFilter.hasActiveFiltersInScope(scopeId)) {
                sheet.AutoFilter._applyFilters(scopeId, {
                    silent: true,
                    emitEvents: false,
                    recordChange: false,
                    recalculate: false
                });
            }
        });
    });
}

function _collectAutoFilterScopes(autoFilter) {
    if (!autoFilter?.getEnabledScopes) return [];
    return autoFilter.getEnabledScopes().map(scope => ({
        scopeId: scope.scopeId,
        ownerType: scope.ownerType,
        ownerId: scope.ownerId,
        range: _clone(scope.range),
        xml: autoFilter.buildXml(scope.scopeId)
    })).filter(item => !!item.xml);
}

/**
 * @returns {Promise<Object>}
 */
export async function getData() {
    const SN = this._SN;

    const fontTable = new IndexTable();
    const fillTable = new IndexTable();
    const borderTable = new IndexTable();
    const alignTable = new IndexTable();
    const styleTable = new IndexTable();
    const rowHeightTable = new IndexTable();
    const colWidthTable = new IndexTable();

    const workbookData = {
        version: JSON_VERSION,
        workbookName: SN.workbookName,
        activeSheet: SN.activeSheet?.name || null,
        calcMode: SN.calcMode === 'manual' ? 'manual' : 'auto',
        readOnly: !!SN.readOnly,
        properties: _clone(SN.properties || {}),
        definedNames: _clone(SN._definedNames || {}),
        styleTable: {
            fonts: [],
            fills: [],
            borders: [],
            aligns: [],
            styles: []
        },
        sizeTable: {
            rowHeights: [],
            colWidths: []
        },
        sheets: []
    };

    for (const sheet of SN.sheets) {
        sheet._init();

        const sheetData = {
            name: sheet.name,
            hidden: !!sheet.hidden,
            showGridLines: sheet.showGridLines !== false,
            showRowColHeaders: sheet.showRowColHeaders !== false,
            showPageBreaks: !!sheet.showPageBreaks,
            outlinePr: _clone(sheet.outlinePr || {}),
            frozenCols: _safeInt(sheet.frozenCols, 0),
            frozenRows: _safeInt(sheet.frozenRows, 0),
            freezeStartRow: _safeInt(sheet._freezeStartRow, 0),
            freezeStartCol: _safeInt(sheet._freezeStartCol, 0),
            defaultColWidth: _safeInt(sheet.defaultColWidth, 72),
            defaultRowHeight: _safeInt(sheet.defaultRowHeight, 21),
            zoom: _normalizeZoom(sheet.zoom),
            activeCell: _clone(_normalizeCellRef(sheet.activeCell, { r: 0, c: 0 })),
            viewStart: _clone(_normalizeCellRef(sheet.viewStart, { r: 0, c: 0 })),
            printSettings: sheet.printSettings ? _clone(sheet.printSettings) : null,
            rowCount: sheet.rows.length,
            colCount: sheet.cols.length,
            cols: [],
            rows: [],
            merges: [],
            drawings: []
        };

        const protection = sheet.protection?.getOptions?.();
        if (protection) sheetData.sheetProtection = _clone(protection);

        sheet.cols.forEach(col => {
            if (!col?._needBuild) return;
            const colData = {
                cIndex: col.cIndex,
                w: colWidthTable.add(col._width)
            };
            if (col._hidden) colData.hidden = true;
            if (col._outlineLevel > 0) colData.ol = col._outlineLevel;
            if (col._collapsed) colData.cl = true;
            if (col._colStyle) colData.s = styleTable.add(col._colStyle);
            sheetData.cols.push(colData);
        });

        sheet.rows.forEach(row => {
            if (!row?._needBuild) return;

            const rowData = {
                rIndex: row.rIndex,
                h: rowHeightTable.add(row._height),
                cells: []
            };
            if (row._hidden) rowData.hidden = true;
            if (row._outlineLevel > 0) rowData.ol = row._outlineLevel;
            if (row._collapsed) rowData.cl = true;
            if (row._rowStyle) rowData.s = styleTable.add(row._rowStyle);

            const tempCells = [];
            row.cells.forEach(cell => {
                if (!cell?._needBuild) return;
                const cellData = { c: cell.cIndex };

                if (typeof cell._editVal === 'string' && cell._editVal.startsWith('=')) {
                    cellData.f = cell._editVal;
                    if (cell._calcVal !== undefined) {
                        const encoded = _encodeCellValue(cell._calcVal, cell.type);
                        cellData.cv = encoded.value;
                        if (encoded.type) cellData.cvt = encoded.type;
                    }
                } else if (cell._editVal !== undefined && cell._editVal !== null && cell._editVal !== '') {
                    const encoded = _encodeCellValue(cell._editVal, cell._type);
                    cellData.v = encoded.value;
                    if (encoded.type) cellData.vt = encoded.type;
                }

                if (cell._style && Object.keys(cell._style).length > 0) {
                    cellData.s = styleTable.add(cell._style);
                    if (cell._style.font) cellData.fo = fontTable.add(cell._style.font);
                    if (cell._style.fill) cellData.fi = fillTable.add(cell._style.fill);
                    if (cell._style.border) cellData.b = borderTable.add(cell._style.border);
                    if (cell._style.alignment) cellData.a = alignTable.add(cell._style.alignment);
                    if (cell._style.numFmt) cellData.fmt = cell._style.numFmt;
                    if (cell._style.protection) cellData.p = _clone(cell._style.protection);
                }

                if (cell.hyperlink) {
                    cellData.link = {
                        t: cell.hyperlink.target ?? null,
                        l: cell.hyperlink.location ?? null
                    };
                    if (cell.hyperlink.tooltip) cellData.link.tt = cell.hyperlink.tooltip;
                }
                if (cell._dataValidation) cellData.dv = _clone(cell._dataValidation);
                if (cell.richText) cellData.rt = _clone(cell.richText);

                tempCells.push(cellData);
            });

            for (let i = 0; i < tempCells.length; i++) {
                const currentCell = tempCells[i];
                let endCol = currentCell.c;
                for (let j = i + 1; j < tempCells.length; j++) {
                    const nextCell = tempCells[j];
                    if (nextCell.c === endCol + 1 && isCellDataEqual(currentCell, nextCell)) {
                        endCol = nextCell.c;
                    } else {
                        break;
                    }
                }
                if (endCol > currentCell.c) {
                    currentCell.e = endCol;
                    i += (endCol - currentCell.c);
                }
                rowData.cells.push(currentCell);
            }

            if (
                rowData.cells.length > 0 ||
                rowData.hidden ||
                rowData.ol !== undefined ||
                rowData.cl ||
                rowData.s !== undefined
            ) {
                sheetData.rows.push(rowData);
            }
        });
        sheet.merges.forEach(merge => {
            const range = _normalizeRange(merge);
            if (!range) return;
            sheetData.merges.push(range);
        });

        const drawings = sheet.Drawing?.getAll?.() || [];
        const initPromises = drawings
            .filter(item => item?._initPromise)
            .map(item => item._initPromise.catch(() => null));
        if (initPromises.length > 0) {
            await Promise.all(initPromises);
        }

        drawings
            .filter(item => item && item.type !== 'slicer')
            .forEach(item => sheetData.drawings.push(_serializeDrawing(item)));

        if (sheet.Sparkline?.groups?.length > 0) {
            sheetData.sparklines = sheet.Sparkline.groups.map(group => {
                const instances = sheet.Sparkline.getAll().filter(inst => inst.group.id === group.id);
                return {
                    type: group.type,
                    displayEmptyCellsAs: group.displayEmptyCellsAs,
                    colors: _clone(group.colors),
                    lineWeight: group.lineWeight,
                    manualMax: group.manualMax,
                    manualMin: group.manualMin,
                    dateAxis: group.dateAxis,
                    markers: group.markers,
                    high: group.high,
                    low: group.low,
                    first: group.first,
                    last: group.last,
                    negative: group.negative,
                    instances: instances.map(inst => ({
                        r: inst.r,
                        c: inst.c,
                        formula: inst.formula
                    }))
                };
            });
        }

        if (sheet.Comment?.map?.size > 0) {
            sheetData.comments = sheet.Comment.getAll().map(comment => ({
                cellRef: comment.cellRef,
                author: comment.author,
                text: comment.text,
                visible: comment.visible,
                width: comment.width,
                height: comment.height,
                marginLeft: comment.marginLeft,
                marginTop: comment.marginTop
            }));
        }

        if (sheet.Table?.size > 0) {
            sheetData.tables = _clone(sheet.Table.toJSON());
        }

        const autoFilters = _collectAutoFilterScopes(sheet.AutoFilter);
        if (autoFilters.length > 0) {
            sheetData.autoFilters = autoFilters;
        }

        const cfRules = _serializeCfRules(sheet);
        if (cfRules.length > 0) {
            sheetData.cfRules = cfRules;
        }

        if (sheet.PivotTable?.size > 0) {
            sheetData.pivotTables = sheet.PivotTable.getAll().map(pt => ({
                name: pt.name,
                cacheId: pt.cacheId,
                xml: pt.toXml()?.pivotTableDefinition || null
            })).filter(item => !!item.xml);
        }

        if (sheet.Slicer?.size > 0) {
            sheetData.slicers = sheet.Slicer.getAll().map(slicer => _serializeSlicer(slicer));
        }

        workbookData.sheets.push(sheetData);
    }

    workbookData.styleTable.fonts = fontTable.getAll();
    workbookData.styleTable.fills = fillTable.getAll();
    workbookData.styleTable.borders = borderTable.getAll();
    workbookData.styleTable.aligns = alignTable.getAll();
    workbookData.styleTable.styles = styleTable.getAll();
    workbookData.sizeTable.rowHeights = rowHeightTable.getAll();
    workbookData.sizeTable.colWidths = colWidthTable.getAll();

    if (SN._pivotCaches?.size > 0) {
        workbookData.pivotCaches = Array.from(SN._pivotCaches.values()).map(cache => _serializePivotCache(cache));
    }

    if (SN._slicerCaches?.size > 0) {
        workbookData.slicerCaches = Array.from(SN._slicerCaches.values()).map(cache => _serializeSlicerCache(cache));
    }

    return workbookData;
}

/**
 * @param {Object} data
 * @returns {boolean}
 */
export function setData(data) {
    const SN = this._SN;

    try {
        if (!data || data.version !== JSON_VERSION || !Array.isArray(data.sheets) || data.sheets.length === 0) {
            SN.Utils.toast(SN.t('core.iO.jsonIO.toast.son'));
            return false;
        }

        const styleTable = {
            fonts: Array.isArray(data.styleTable?.fonts) ? data.styleTable.fonts : [],
            fills: Array.isArray(data.styleTable?.fills) ? data.styleTable.fills : [],
            borders: Array.isArray(data.styleTable?.borders) ? data.styleTable.borders : [],
            aligns: Array.isArray(data.styleTable?.aligns) ? data.styleTable.aligns : [],
            styles: Array.isArray(data.styleTable?.styles) ? data.styleTable.styles : []
        };

        const sizeTable = {
            rowHeights: Array.isArray(data.sizeTable?.rowHeights) ? data.sizeTable.rowHeights : [],
            colWidths: Array.isArray(data.sizeTable?.colWidths) ? data.sizeTable.colWidths : []
        };

        if (typeof data.workbookName === 'string' && data.workbookName.trim()) {
            SN._workbookName = data.workbookName.trim();
            const fileNameDom = SN.containerDom?.querySelector?.(".sn-file-name");
            if (fileNameDom) fileNameDom.textContent = SN._workbookName;
        }

        SN.calcMode = data.calcMode === 'manual' ? 'manual' : 'auto';
        SN.readOnly = !!data.readOnly;
        SN.properties = _isObject(data.properties) ? _clone(data.properties) : {};
        SN._definedNames = _isObject(data.definedNames) ? _clone(data.definedNames) : {};

        SN._sharedFormula = {};
        SN._pivotCaches = new Map();
        SN._slicerCaches = new Map();
        SN._slicerXmlNodes = [];
        SN._excelSlicerCacheMeta = new Map();

        SN.sheets = [];

        const pendingAutoFilters = [];
        const pendingPivotTables = [];
        const pendingSlicers = [];

        data.sheets.forEach((sheetData, index) => {
            const fallbackName = `Sheet${index + 1}`;
            const sheetName = typeof sheetData?.name === 'string' && sheetData.name.trim()
                ? sheetData.name.trim()
                : fallbackName;

            let sheet = SN.addSheet(sheetName);
            if (!sheet && sheetName !== fallbackName) {
                sheet = SN.addSheet(fallbackName);
            }
            if (!sheet) return;

            _resetSheetForImport(sheet);

            sheet._hidden = !!sheetData.hidden;
            sheet.showGridLines = sheetData.showGridLines !== false;
            sheet.showRowColHeaders = sheetData.showRowColHeaders !== false;
            sheet.showPageBreaks = !!sheetData.showPageBreaks;
            sheet.outlinePr = {
                applyStyles: false,
                summaryBelow: true,
                summaryRight: true,
                showOutlineSymbols: true,
                ...(sheetData.outlinePr || {})
            };

            const defaultColWidth = _safeNumber(sheetData.defaultColWidth, null);
            if (defaultColWidth && defaultColWidth > 0) {
                sheet.defaultColWidth = Math.round(defaultColWidth);
            }
            const defaultRowHeight = _safeNumber(sheetData.defaultRowHeight, null);
            if (defaultRowHeight && defaultRowHeight > 0) {
                sheet.defaultRowHeight = Math.round(defaultRowHeight);
            }

            sheet._frozenCols = Math.max(0, _safeInt(sheetData.frozenCols, 0));
            sheet._frozenRows = Math.max(0, _safeInt(sheetData.frozenRows, 0));
            sheet._freezeStartRow = Math.max(0, _safeInt(sheetData.freezeStartRow, 0));
            sheet._freezeStartCol = Math.max(0, _safeInt(sheetData.freezeStartCol, 0));

            sheet._zoom = _normalizeZoom(sheetData.zoom);
            const activeCell = _normalizeCellRef(sheetData.activeCell, { r: 0, c: 0 });
            const viewStart = _normalizeCellRef(sheetData.viewStart, { r: 0, c: 0 });
            sheet._activeCell = activeCell;
            sheet.vi = { rowArr: [viewStart.r], colArr: [viewStart.c] };

            if (sheetData.printSettings) {
                sheet.printSettings = _clone(sheetData.printSettings);
            } else {
                sheet.printSettings = SN.Print?.createDefaultSettings?.() || null;
            }

            const { rowCount, colCount } = _resolveSheetSize(sheetData);
            _ensureSheetSize(sheet, rowCount, colCount);

            if (Array.isArray(sheetData.cols)) {
                sheetData.cols.forEach(colData => {
                    const cIndex = Math.max(0, _safeInt(colData?.cIndex, 0));
                    const col = sheet.getCol(cIndex);

                    const widthByTable = colData?.w !== undefined ? sizeTable.colWidths[colData.w] : undefined;
                    const width = _safeNumber(widthByTable, null);
                    if (width && width > 0) col._width = Math.round(width);

                    if (colData?.hidden !== undefined) col._hidden = !!colData.hidden;
                    if (colData?.ol !== undefined) {
                        col._outlineLevel = Math.max(0, _safeInt(colData.ol, 0));
                    }
                    if (colData?.cl !== undefined) {
                        col._collapsed = !!colData.cl;
                    }

                    if (colData?.s !== undefined && styleTable.styles[colData.s]) {
                        col._colStyle = _clone(styleTable.styles[colData.s]);
                    }
                });
            }
            if (Array.isArray(sheetData.rows)) {
                sheetData.rows.forEach(rowData => {
                    const rIndex = Math.max(0, _safeInt(rowData?.rIndex, 0));
                    const row = sheet.getRow(rIndex);

                    const heightByTable = rowData?.h !== undefined ? sizeTable.rowHeights[rowData.h] : undefined;
                    const height = _safeNumber(heightByTable, null);
                    if (height && height > 0) row._height = Math.round(height);

                    if (rowData?.hidden !== undefined) row._hidden = !!rowData.hidden;
                    if (rowData?.ol !== undefined) {
                        row._outlineLevel = Math.max(0, _safeInt(rowData.ol, 0));
                    }
                    if (rowData?.cl !== undefined) {
                        row._collapsed = !!rowData.cl;
                    }

                    if (rowData?.s !== undefined && styleTable.styles[rowData.s]) {
                        row._rowStyle = _clone(styleTable.styles[rowData.s]);
                    }

                    if (!Array.isArray(rowData?.cells)) return;

                    rowData.cells.forEach(cellData => {
                        const startCol = Math.max(0, _safeInt(cellData?.c, 0));
                        const endCol = Math.max(startCol, _safeInt(cellData?.e, startCol));
                        const template = _createDecodedCellTemplate(cellData, styleTable);

                        for (let cIndex = startCol; cIndex <= endCol; cIndex++) {
                            const cell = sheet.getCell(rIndex, cIndex);
                            _applyDecodedCellTemplate(cell, template);
                        }
                    });
                });
            }

            sheet.merges = [];
            if (Array.isArray(sheetData.merges)) {
                sheetData.merges.forEach(mergeData => {
                    const range = _normalizeRange(mergeData);
                    if (!range) return;
                    sheet.merges.push(range);
                    sheet.eachCells(range, (r, c) => {
                        const cell = sheet.getCell(r, c);
                        cell.isMerged = true;
                        cell.master = { r: range.s.r, c: range.s.c };
                    });
                });
            }

            if (Array.isArray(sheetData.drawings)) {
                sheetData.drawings.forEach(drawingData => _applyDrawingData(sheet, drawingData));
            }

            if (Array.isArray(sheetData.sparklines) && sheet.Sparkline) {
                sheet.Sparkline.groups = [];
                sheet.Sparkline.map = new Map();

                sheetData.sparklines.forEach(groupData => {
                    const group = {
                        id: sheet.Sparkline.groups.length,
                        type: groupData.type || 'line',
                        displayEmptyCellsAs: groupData.displayEmptyCellsAs || 'gap',
                        colors: groupData.colors || {
                            series: '#376092',
                            negative: '#D00000',
                            axis: '#000000',
                            markers: '#D00000',
                            first: '#D00000',
                            last: '#D00000',
                            high: '#D00000',
                            low: '#D00000'
                        },
                        lineWeight: groupData.lineWeight ?? 0.75,
                        manualMax: groupData.manualMax ?? null,
                        manualMin: groupData.manualMin ?? null,
                        dateAxis: groupData.dateAxis ?? false,
                        markers: groupData.markers ?? false,
                        high: groupData.high ?? false,
                        low: groupData.low ?? false,
                        first: groupData.first ?? false,
                        last: groupData.last ?? false,
                        negative: groupData.negative ?? false
                    };
                    sheet.Sparkline.groups.push(group);

                    const instances = Array.isArray(groupData.instances) ? groupData.instances : [];
                    instances.forEach(instanceData => {
                        const r = Math.max(0, _safeInt(instanceData?.r, 0));
                        const c = Math.max(0, _safeInt(instanceData?.c, 0));
                        const key = `${r}:${c}`;
                        const instance = {
                            key,
                            r,
                            c,
                            sqref: sheet.SN.Utils.cellNumToStr({ r, c }),
                            formula: instanceData?.formula || '',
                            group,
                            _dataCache: null,
                            _renderCache: null
                        };
                        sheet.Sparkline.map.set(key, instance);
                    });
                });
            }

            if (Array.isArray(sheetData.comments) && sheet.Comment) {
                sheetData.comments.forEach(commentData => {
                    if (!commentData?.cellRef) return;
                    sheet.Comment.add({
                        cellRef: commentData.cellRef,
                        author: commentData.author,
                        text: commentData.text,
                        visible: !!commentData.visible,
                        width: _safeNumber(commentData.width, 144) ?? 144,
                        height: _safeNumber(commentData.height, 72) ?? 72,
                        marginLeft: _safeNumber(commentData.marginLeft, 107) ?? 107,
                        marginTop: _safeNumber(commentData.marginTop, 7) ?? 7
                    });
                });
            }

            if (Array.isArray(sheetData.tables) && sheet.Table) {
                sheet.Table.fromJSON(_clone(sheetData.tables));
            }

            if (Array.isArray(sheetData.cfRules) && sheet.CF) {
                _restoreCfRules(sheet, sheetData.cfRules);
            }

            if (sheetData.sheetProtection) {
                sheet.protection.enable(_clone(sheetData.sheetProtection));
            }

            if (Array.isArray(sheetData.autoFilters)) {
                pendingAutoFilters.push({ sheet, scopes: _clone(sheetData.autoFilters) });
            }

            if (Array.isArray(sheetData.pivotTables)) {
                pendingPivotTables.push({ sheet, pivotTables: _clone(sheetData.pivotTables) });
            }

            if (Array.isArray(sheetData.slicers)) {
                pendingSlicers.push({ sheet, slicers: _clone(sheetData.slicers) });
            }
        });

        _restorePivotCaches(SN, data.pivotCaches);
        _restoreSlicerCaches(SN, data.slicerCaches);
        _restorePivotTables(SN, pendingPivotTables);
        _restoreAutoFilters(pendingAutoFilters);
        _restoreSlicers(SN, pendingSlicers);

        SN.sheets.forEach(sheet => {
            if (!sheet?.PivotTable || sheet.PivotTable.size === 0) return;
            sheet.PivotTable.getAll().forEach(pt => {
                const isEmpty = pt.rowFields.length === 0 &&
                    pt.colFields.length === 0 &&
                    pt.dataFields.length === 0 &&
                    pt.pageFields.length === 0;
                if (isEmpty) return;
                pt._needsRecalculate = true;
                pt.calculate();
                pt.render({ silent: true });
            });
        });

        const requestedActiveSheet = typeof data.activeSheet === 'string' ? SN.getSheet(data.activeSheet) : null;
        const activeSheet = requestedActiveSheet || SN.sheets.find(sheet => !sheet.hidden) || SN.sheets[0] || null;
        if (!activeSheet) return false;
        SN.activeSheet = activeSheet;
        if (activeSheet._updView) activeSheet._updView();

        if (SN.DependencyGraph) {
            SN.DependencyGraph.clear();
            SN.DependencyGraph.rebuildAll();
            if (SN.calcMode !== 'manual') {
                SN.DependencyGraph.recalculateVolatile();
            }
        }

        SN.UndoRedo?.clear?.();
        return true;
    } catch (error) {
        console.error("JSON setData failed:", error);
        SN.Utils.toast(SN.t('core.iO.jsonIO.toast.loadFailed', { message: error.message }));
        return false;
    }
}

/**
 * @returns {Promise<void>}
 */
export async function exportJson() {
    const SN = this._SN;
    console.time("JSON export");
    const workbookData = await getData.call(this);
    const jsonString = JSON.stringify(workbookData);
    const blob = new Blob([jsonString], { type: 'application/json' });
    saveAs(blob, `${SN.workbookName}.json`);
    console.timeEnd("JSON export");
}

/**
 * @param {File} file
 * @returns {Promise<void>}
 */
export async function importJson(file) {
    const SN = this._SN;
    try {
        console.time("JSON import");
        const text = await file.text();
        const data = JSON.parse(text);
        const success = setData.call(this, data);
        if (success) {
            console.timeEnd("JSON import");
            SN.Utils.toast(SN.t('core.iO.jsonIO.toast.json'));
        } else {
            console.timeEnd("JSON import");
        }
    } catch (error) {
        console.error("JSON import failed:", error);
        SN.Utils.toast(SN.t('core.iO.jsonIO.toast.importFailed', { message: error.message }));
    }
}
