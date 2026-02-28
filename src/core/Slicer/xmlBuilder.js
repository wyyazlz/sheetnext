const CUSTOM_SLICER_NS = 'http://sheetnext/slicer';
const EXCEL_SLICER_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const MARKUP_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const XR10_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10';
const SLICER_DRAWING_NS = 'http://schemas.microsoft.com/office/drawing/2010/slicer';

function ensureArray(val) {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}

function toNumber(value, fallback = 0) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
}

function ensureOverride(_Xml, partName, contentType) {
    const overrides = _Xml.obj['[Content_Types].xml']?.Types?.Override;
    if (!overrides) return;
    const exists = overrides.some(item => item._$PartName === partName);
    if (!exists) {
        overrides.push({
            '_$PartName': partName,
            '_$ContentType': contentType
        });
    }
}

function getNextRid(relationships) {
    let max = 0;
    ensureArray(relationships).forEach(rel => {
        const match = String(rel?._$Id || '').match(/^rId(\d+)$/);
        if (!match) return;
        const num = Number(match[1]);
        if (num > max) max = num;
    });
    return max + 1;
}

function collectSlicers(SN) {
    const slicers = [];
    SN.sheets.forEach((sheet, sheetIndex) => {
        if (!sheet.Slicer || sheet.Slicer.size === 0) return;
        sheet.Slicer.forEach(slicer => {
            slicers.push({ sheet, sheetIndex, slicer });
        });
    });
    return slicers;
}

function buildCacheItems(cache) {
    const items = cache.getItems({ force: true });
    const list = items.map(item => ({
        '_$key': item.key,
        '_$value': item.value === null || item.value === undefined ? '' : String(item.value),
        '_$text': item.text ?? ''
    }));
    return list.length > 0 ? { 'sn:item': list } : undefined;
}

function buildCacheNode(cache) {
    const node = {
        '_$id': String(cache.cacheId),
        '_$type': cache.sourceType,
        '_$sheet': cache.sourceSheet ?? '',
        '_$name': cache.sourceName ?? '',
        '_$fieldIndex': String(cache.fieldIndex ?? 0),
        '_$fieldName': cache.fieldName ?? ''
    };
    const itemsNode = buildCacheItems(cache);
    if (itemsNode) node['sn:items'] = itemsNode;
    return node;
}

function buildSlicerAnchor(slicer) {
    const drawing = slicer.Drawing;
    if (!drawing) return null;
    return {
        '_$startRow': String(drawing.startCell?.r ?? 0),
        '_$startCol': String(drawing.startCell?.c ?? 0),
        '_$offsetX': String(drawing.offsetX ?? 0),
        '_$offsetY': String(drawing.offsetY ?? 0),
        '_$width': String(drawing.width ?? 0),
        '_$height': String(drawing.height ?? 0)
    };
}

function buildSlicerSelection(slicer) {
    const keys = Array.from(slicer.selectedKeys || []);
    if (keys.length === 0) return null;
    return {
        'sn:item': keys.map(key => ({ '_$key': key }))
    };
}

function buildSlicerNode(slicer) {
    const node = {
        '_$id': slicer.id,
        '_$name': slicer.name,
        '_$caption': slicer.caption ?? '',
        '_$cacheId': String(slicer.cacheId ?? ''),
        '_$sheet': slicer.sheet?.name ?? slicer.sourceSheet ?? '',
        '_$sourceType': slicer.sourceType ?? '',
        '_$sourceName': slicer.sourceName ?? '',
        '_$fieldIndex': String(slicer.fieldIndex ?? 0),
        '_$fieldName': slicer.fieldName ?? '',
        '_$columns': String(slicer.columns ?? 1),
        '_$buttonHeight': String(slicer.buttonHeight ?? 22),
        '_$showHeader': slicer.showHeader ? '1' : '0',
        '_$multiSelect': slicer.multiSelect ? '1' : '0',
        '_$styleName': slicer.styleName ?? 'SlicerStyleLight1'
    };

    const selection = buildSlicerSelection(slicer);
    if (selection) node['sn:selection'] = selection;

    const anchor = buildSlicerAnchor(slicer);
    if (anchor) node['sn:anchor'] = anchor;

    return node;
}

function exportCustomSlicers(SN, _Xml, slicerEntries, caches) {
    if (caches.length === 0 && slicerEntries.length === 0) return;

    const cacheNodes = caches.map(buildCacheNode);
    const slicerNodes = slicerEntries.map(({ slicer }) => buildSlicerNode(slicer));

    _Xml.obj['xl/snSlicers.xml'] = {
        '?xml': { '_$version': '1.0', '_$encoding': 'UTF-8', '_$standalone': 'yes' },
        'sn:slicers': {
            '_$xmlns:sn': CUSTOM_SLICER_NS,
            'sn:caches': cacheNodes.length > 0 ? { 'sn:cache': cacheNodes } : undefined,
            'sn:slicer': slicerNodes.length > 0 ? slicerNodes : undefined
        }
    };

    ensureOverride(_Xml, '/xl/snSlicers.xml', 'application/vnd.sheetnext.Slicer+xml');
}

function buildExportedTableIdLookup(sheet, sheetIndex) {
    const lookup = new Map();
    if (!sheet?.Table || sheet.Table.size === 0) return lookup;

    let tableCount = 0;
    sheet.Table.forEach(table => {
        tableCount += 1;
        const tableIndex = sheetIndex * 100 + tableCount;
        const exportedTableId = tableIndex + 1;

        lookup.set(table.id, exportedTableId);
        if (table.name) lookup.set(`name:${table.name}`, exportedTableId);
        if (table.displayName) lookup.set(`name:${table.displayName}`, exportedTableId);
    });

    return lookup;
}

function resolveStandardTableId(cache, slicer, tableLookup) {
    if (!tableLookup || tableLookup.size === 0) return 1;

    if (cache?.sourceId && tableLookup.has(cache.sourceId)) {
        return tableLookup.get(cache.sourceId);
    }

    if (slicer?.sourceId && tableLookup.has(slicer.sourceId)) {
        return tableLookup.get(slicer.sourceId);
    }

    if (cache?.sourceName && tableLookup.has(`name:${cache.sourceName}`)) {
        return tableLookup.get(`name:${cache.sourceName}`);
    }

    if (slicer?.sourceName && tableLookup.has(`name:${slicer.sourceName}`)) {
        return tableLookup.get(`name:${slicer.sourceName}`);
    }

    const first = tableLookup.values().next();
    return first.done ? 1 : first.value;
}

function buildStandardCacheXml(cacheInfo) {
    return {
        '?xml': { '_$version': '1.0', '_$encoding': 'UTF-8', '_$standalone': 'yes' },
        slicerCacheDefinition: {
            '_$xmlns': EXCEL_SLICER_NS,
            '_$xmlns:mc': MARKUP_NS,
            '_$mc:Ignorable': 'x xr10',
            '_$xmlns:x': SPREADSHEET_NS,
            '_$xmlns:xr10': XR10_NS,
            '_$name': cacheInfo.cacheName,
            '_$sourceName': cacheInfo.sourceName,
            extLst: {
                'x:ext': {
                    '_$uri': '{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}',
                    '_$xmlns:x15': X15_NS,
                    'x15:tableSlicerCache': {
                        '_$tableId': String(cacheInfo.tableId),
                        '_$column': String(cacheInfo.column)
                    }
                }
            }
        }
    };
}

function buildStandardSlicerXml(slicerInfo, cacheName) {
    return {
        '?xml': { '_$version': '1.0', '_$encoding': 'UTF-8', '_$standalone': 'yes' },
        slicers: {
            '_$xmlns': EXCEL_SLICER_NS,
            '_$xmlns:mc': MARKUP_NS,
            '_$mc:Ignorable': 'x xr10',
            '_$xmlns:x': SPREADSHEET_NS,
            '_$xmlns:xr10': XR10_NS,
            slicer: {
                '_$name': slicerInfo.name,
                '_$cache': cacheName,
                '_$caption': slicerInfo.caption,
                '_$rowHeight': String(slicerInfo.rowHeight),
                '_$columnCount': String(slicerInfo.columnCount),
                '_$showCaption': slicerInfo.showCaption,
                '_$style': slicerInfo.styleName
            }
        }
    };
}

function appendWorkbookSlicerCacheExt(workbook, cacheRids) {
    if (!workbook || cacheRids.length === 0) return;

    const extArr = ensureArray(workbook.extLst?.ext);
    let targetExt = extArr.find(ext => ext?.['x15:slicerCaches']);

    if (!targetExt) {
        targetExt = {
            '_$uri': '{46BE6895-7355-4a93-B00E-2C351335B9C9}',
            '_$xmlns:x15': X15_NS,
            'x15:slicerCaches': {
                '_$xmlns:x14': X14_NS,
                'x14:slicerCache': []
            }
        };
        extArr.push(targetExt);
    }

    targetExt['x15:slicerCaches'] = {
        '_$xmlns:x14': X14_NS,
        'x14:slicerCache': cacheRids.map(rid => ({ '_$r:id': rid }))
    };

    workbook.extLst = { ext: extArr };
}

function appendSheetSlicerExt(worksheet, slicerRids) {
    if (!worksheet || slicerRids.length === 0) return;

    const extArr = ensureArray(worksheet.extLst?.ext);
    let targetExt = extArr.find(ext => ext?.['x14:slicerList']);

    if (!targetExt) {
        targetExt = {
            '_$uri': '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}',
            '_$xmlns:x15': X15_NS,
            'x14:slicerList': {
                '_$xmlns:x14': X14_NS,
                'x14:slicer': []
            }
        };
        extArr.push(targetExt);
    }

    targetExt['x14:slicerList'] = {
        '_$xmlns:x14': X14_NS,
        'x14:slicer': slicerRids.map(rid => ({ '_$r:id': rid }))
    };

    worksheet.extLst = { ext: extArr };
}

function exportStandardSlicers(SN, _Xml, slicerEntries, caches) {
    if (slicerEntries.length === 0) return;

    const workbook = _Xml.obj['xl/workbook.xml']?.workbook;
    const workbookRels = ensureArray(_Xml.relationship);
    if (!SN._definedNames) SN._definedNames = {};
    const tableLookupBySheet = new Map();

    SN.sheets.forEach((sheet, sheetIndex) => {
        tableLookupBySheet.set(sheet.name, buildExportedTableIdLookup(sheet, sheetIndex));
    });

    const cacheMetaById = new Map();
    const cacheRidList = [];
    let cacheFileIndex = 0;

    caches.forEach(cache => {
        const linked = slicerEntries.find(item => item.slicer.cacheId === cache.cacheId);
        const sourceSheetName = cache.sourceSheet || linked?.sheet?.name;
        const tableLookup = tableLookupBySheet.get(sourceSheetName);
        const tableId = resolveStandardTableId(cache, linked?.slicer, tableLookup);
        const column = Math.max(1, toNumber(cache.fieldIndex, 0) + 1);

        const cacheFileName = `slicerCache${++cacheFileIndex}.xml`;
        const cacheRid = `rId${getNextRid(workbookRels)}`;
        const cacheNameRaw = cache._excelCacheName || `SlicerCache_${linked?.slicer?.name || cache.fieldName || cache.cacheId}`;
        const cacheName = String(cacheNameRaw).replace(/\s+/g, '_');
        const sourceName = cache.fieldName || linked?.slicer?.fieldName || '';

        workbookRels.push({
            '_$Id': cacheRid,
            '_$Type': 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
            '_$Target': `slicerCaches/${cacheFileName}`
        });

        cacheRidList.push(cacheRid);
        cacheMetaById.set(cache.cacheId, { cacheName });
        SN._definedNames[cacheName] = SN._definedNames[cacheName] || '#N/A';

        _Xml.obj[`xl/slicerCaches/${cacheFileName}`] = buildStandardCacheXml({
            cacheName,
            sourceName,
            tableId,
            column
        });

        ensureOverride(_Xml, `/xl/slicerCaches/${cacheFileName}`, 'application/vnd.ms-excel.slicerCache+xml');
    });

    appendWorkbookSlicerCacheExt(workbook, cacheRidList);

    let slicerFileIndex = 0;
    SN.sheets.forEach((sheet, sheetIndex) => {
        if (!sheet?.Slicer || sheet.Slicer.size === 0) return;

        const sheetFile = `sheet${sheetIndex + 1}.xml`;
        const worksheet = _Xml.obj[`xl/worksheets/${sheetFile}`]?.worksheet;
        const sheetRels = ensureArray(_Xml.getRelsByTarget(`worksheets/${sheetFile}`, true));
        const sheetSlicerRids = [];

        sheet.Slicer.forEach(slicer => {
            const cacheMeta = cacheMetaById.get(slicer.cacheId);
            if (!cacheMeta) return;

            const slicerFileName = `slicer${++slicerFileIndex}.xml`;
            const slicerRid = `rId${getNextRid(sheetRels)}`;

            sheetRels.push({
                '_$Id': slicerRid,
                '_$Type': 'http://schemas.microsoft.com/office/2007/relationships/slicer',
                '_$Target': `../slicers/${slicerFileName}`
            });

            sheetSlicerRids.push(slicerRid);

            const rowHeight = Math.max(1, Math.round(SN.Utils.pxToEmu(toNumber(slicer.buttonHeight, 22))));
            _Xml.obj[`xl/slicers/${slicerFileName}`] = buildStandardSlicerXml({
                name: slicer.name || slicer.id,
                caption: slicer.caption || slicer.fieldName || slicer.name || slicer.id,
                rowHeight,
                columnCount: Math.max(1, toNumber(slicer.columns, 1)),
                showCaption: slicer.showHeader ? '1' : '0',
                styleName: slicer.styleName || 'SlicerStyleLight1'
            }, cacheMeta.cacheName);

            ensureOverride(_Xml, `/xl/slicers/${slicerFileName}`, 'application/vnd.ms-excel.slicer+xml');
        });

        appendSheetSlicerExt(worksheet, sheetSlicerRids);
    });
}

export function buildSlicerGraphicFrame(index, slicerName) {
    const safeName = slicerName || `Slicer_${index + 1}`;
    const graphicFrame = {
        'xdr:nvGraphicFramePr': {
            'xdr:cNvPr': {
                '_$id': String(index + 2),
                '_$name': safeName
            },
            'xdr:cNvGraphicFramePr': ''
        },
        'xdr:xfrm': {
            'a:off': { '_$x': '0', '_$y': '0' },
            'a:ext': { '_$cx': '0', '_$cy': '0' }
        },
        'a:graphic': {
            'a:graphicData': {
                '_$uri': SLICER_DRAWING_NS,
                'sle:slicer': {
                    '_$xmlns:sle': SLICER_DRAWING_NS,
                    '_$name': safeName
                }
            }
        },
        '_$macro': ''
    };

    return {
        'mc:Choice': {
            '_$Requires': 'sle15',
            'xdr:graphicFrame': graphicFrame
        },
        '_$xmlns:mc': MARKUP_NS,
        '_$xmlns:sle15': 'http://schemas.microsoft.com/office/drawing/2012/slicer'
    };
}

export function exportSlicers(SN, _Xml) {
    const slicerEntries = collectSlicers(SN);
    const cacheIds = Array.from(new Set(
        slicerEntries
            .map(item => item.slicer?.cacheId)
            .filter(id => id !== null && id !== undefined)
    ));
    const caches = cacheIds
        .map(id => SN?._slicerCaches?.get(id))
        .filter(Boolean);

    if (slicerEntries.length === 0 && caches.length === 0) return;

    exportCustomSlicers(SN, _Xml, slicerEntries, caches);
    exportStandardSlicers(SN, _Xml, slicerEntries, caches);
}
