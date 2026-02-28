import { parseSlicerKey, formatSlicerValue } from './helpers.js';
import SlicerCache from './SlicerCache.js';
import SlicerItem from './SlicerItem.js';

function ensureArray(val) {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}

function toNumber(value, fallback = 0) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
}

function normalizeTarget(target) {
    if (!target) return null;
    if (target.startsWith('xl/')) return target;
    if (target.startsWith('../')) return `xl/${target.substring(3)}`;
    return `xl/${target}`;
}

function parseCacheItems(cacheNode) {
    const itemsNode = cacheNode?.['sn:items']?.['sn:item'];
    const items = [];
    ensureArray(itemsNode).forEach(node => {
        const key = node?._$key;
        if (!key) return;
        const value = parseSlicerKey(key);
        items.push({
            key,
            value,
            text: node?._$text ?? formatSlicerValue(value),
            count: Number(node?._$count ?? 0)
        });
    });
    return items;
}

function resolveGraphicFrame(anchor, utils) {
    if (anchor?.['xdr:graphicFrame']) return anchor['xdr:graphicFrame'];
    const alt = anchor?.['mc:AlternateContent'];
    if (!alt) return null;
    const choices = utils.objToArr(alt['mc:Choice']);
    for (const choice of choices) {
        if (choice?.['xdr:graphicFrame']) return choice['xdr:graphicFrame'];
    }
    return null;
}

function calcSpanSize(sheet, startIndex, endIndex, startOffset, endOffset, axis) {
    if (!Number.isInteger(startIndex) || !Number.isInteger(endIndex) || endIndex < startIndex) return 0;
    const getSize = axis === 'row'
        ? (index) => sheet.getRow(index).height
        : (index) => sheet.getCol(index).width;

    let total = 0;
    for (let i = startIndex; i <= endIndex; i++) total += getSize(i);

    const endSize = getSize(endIndex);
    const size = total - startOffset - (endSize - endOffset);
    return size > 0 ? size : 0;
}

function parseStandardSlicerAnchors(sheet, _Xml) {
    const map = new Map();
    const ws = sheet._xmlObj;
    const drawingRid = ws?.drawing?.['_$r:id'];
    if (!drawingRid) return map;

    const drawingPath = sheet.SN.Xml.getWsRelsFileTarget(sheet.rId, drawingRid);
    if (!drawingPath) return map;

    const drawRoot = _Xml.obj[drawingPath]?.['xdr:wsDr'];
    if (!drawRoot) return map;

    const oneCellAnchors = ensureArray(drawRoot['xdr:oneCellAnchor']);
    const twoCellAnchors = ensureArray(drawRoot['xdr:twoCellAnchor']);
    const anchors = oneCellAnchors.concat(twoCellAnchors);

    anchors.forEach(anchor => {
        const graphicFrame = resolveGraphicFrame(anchor, sheet.SN.Utils);
        if (!graphicFrame) return;

        const graphicData = graphicFrame['a:graphic']?.['a:graphicData'];
        const uri = String(graphicData?._$uri || '').toLowerCase();
        if (!uri.includes('/slicer')) return;

        const slicerNode = graphicData?.['sle:slicer'] || graphicData?.['sle15:slicer'];
        const slicerName = slicerNode?._$name
            || graphicFrame?.['xdr:nvGraphicFramePr']?.['xdr:cNvPr']?._$name;
        if (!slicerName) return;

        const from = anchor?.['xdr:from'];
        const to = anchor?.['xdr:to'];
        if (!from || !to) return;

        const startRow = toNumber(from['xdr:row'], 0);
        const startCol = toNumber(from['xdr:col'], 0);
        const endRow = toNumber(to['xdr:row'], startRow);
        const endCol = toNumber(to['xdr:col'], startCol);
        const offsetX = sheet.Utils.emuToPx(toNumber(from['xdr:colOff'], 0));
        const offsetY = sheet.Utils.emuToPx(toNumber(from['xdr:rowOff'], 0));
        const endOffsetX = sheet.Utils.emuToPx(toNumber(to['xdr:colOff'], 0));
        const endOffsetY = sheet.Utils.emuToPx(toNumber(to['xdr:rowOff'], 0));

        const width = calcSpanSize(sheet, startCol, endCol, offsetX, endOffsetX, 'col');
        const height = calcSpanSize(sheet, startRow, endRow, offsetY, endOffsetY, 'row');

        map.set(slicerName, {
            startRow,
            startCol,
            offsetX,
            offsetY,
            width,
            height
        });
    });

    return map;
}

function resolveTableById(sheet, tableId) {
    if (!Number.isFinite(tableId)) return null;
    const tableKey = `table_${tableId}`;

    const current = sheet.Table?.get(tableKey);
    if (current) return current;

    for (const otherSheet of (sheet.SN.sheets || [])) {
        if (otherSheet === sheet) continue;
        const table = otherSheet.Table?.get(tableKey);
        if (table) return table;
    }
    return null;
}

function resolveFieldIndex(table, cacheMeta, slicerNode) {
    const sourceName = cacheMeta?.sourceName || slicerNode?._$name || '';
    const columns = table?.columns || [];

    if (sourceName && columns.length > 0) {
        const indexByName = columns.findIndex(col => col?.name === sourceName);
        if (indexByName >= 0) return indexByName;
    }

    const rawColumn = Number(cacheMeta?.column);
    if (!Number.isFinite(rawColumn)) return 0;

    if (columns.length > 0) {
        if (rawColumn >= 0 && rawColumn < columns.length) {
            if (sourceName && columns[rawColumn]?.name !== sourceName && rawColumn > 0 && columns[rawColumn - 1]?.name === sourceName) {
                return rawColumn - 1;
            }
            return rawColumn;
        }
        if (rawColumn > 0 && rawColumn - 1 < columns.length) return rawColumn - 1;
    }

    return rawColumn > 0 ? rawColumn - 1 : 0;
}

function ensureStandardCache(sheet, cacheMeta, slicerNode) {
    const SN = sheet.SN;
    if (!SN._slicerCaches) SN._slicerCaches = new Map();

    const cacheId = cacheMeta?.cacheId;
    if (Number.isFinite(cacheId) && SN._slicerCaches.has(cacheId)) {
        const existing = SN._slicerCaches.get(cacheId);
        if (cacheMeta?.cacheName && !existing._excelCacheName) existing._excelCacheName = cacheMeta.cacheName;
        return existing;
    }

    const table = resolveTableById(sheet, Number(cacheMeta?.tableId));
    const fieldIndex = resolveFieldIndex(table, cacheMeta, slicerNode);
    const fieldName = table?.columns?.[fieldIndex]?.name
        || cacheMeta?.sourceName
        || slicerNode?._$name
        || '';

    const cache = new SlicerCache(SN, {
        cacheId,
        sourceType: 'table',
        sourceSheet: table?.sheet?.name || sheet.name,
        sourceName: table?.name || table?.displayName || null,
        sourceId: table?.id || (Number.isFinite(Number(cacheMeta?.tableId)) ? `table_${Number(cacheMeta.tableId)}` : null),
        fieldIndex,
        fieldName
    });
    cache._excelCacheName = cacheMeta?.cacheName || null;

    SN._slicerCaches.set(cache.cacheId, cache);
    return cache;
}

function parseStandardSlicerCachesMeta(SN, _Xml) {
    const workbook = _Xml.obj['xl/workbook.xml']?.workbook;
    const extNodes = ensureArray(workbook?.extLst?.ext);
    const cacheRidList = [];

    extNodes.forEach(ext => {
        const slicerCachesNode = ext?.['x15:slicerCaches'];
        if (!slicerCachesNode) return;
        ensureArray(slicerCachesNode?.['x14:slicerCache']).forEach(cacheRef => {
            const rid = cacheRef?.['_$r:id'];
            if (rid) cacheRidList.push(rid);
        });
    });

    const rels = ensureArray(_Xml.relationship);
    const existingIds = SN._slicerCaches ? Array.from(SN._slicerCaches.keys()) : [];
    let nextCacheId = Math.max(0, ...existingIds) + 1;

    SN._excelSlicerCacheMeta = new Map();

    cacheRidList.forEach(rid => {
        const rel = rels.find(item => item?._$Id === rid);
        const target = normalizeTarget(rel?._$Target);
        const cacheDef = target ? _Xml.obj[target]?.slicerCacheDefinition : null;
        if (!cacheDef) return;

        let tableSlicerCache = null;
        const extList = ensureArray(cacheDef?.extLst?.['x:ext']);
        for (const ext of extList) {
            if (ext?.['x15:tableSlicerCache']) {
                tableSlicerCache = ext['x15:tableSlicerCache'];
                break;
            }
        }

        const cacheName = cacheDef?._$name || cacheDef?._$sourceName || rid;
        SN._excelSlicerCacheMeta.set(cacheName, {
            cacheId: nextCacheId++,
            cacheName,
            sourceName: cacheDef?._$sourceName || '',
            tableId: toNumber(tableSlicerCache?._$tableId, NaN),
            column: toNumber(tableSlicerCache?._$column, NaN)
        });
    });
}

function parseCustomSheetSlicers(sheet) {
    const slicerNodes = sheet.SN._slicerXmlNodes || [];
    let parsedCount = 0;

    slicerNodes.forEach(node => {
        if (node?._$sheet !== sheet.name) return;

        const cacheId = Number(node?._$cacheId);
        const cache = sheet.SN._slicerCaches?.get(cacheId) || new SlicerCache(sheet.SN, {
            cacheId,
            sourceType: node?._$sourceType,
            sourceSheet: node?._$sheet,
            sourceName: node?._$sourceName,
            fieldIndex: Number(node?._$fieldIndex ?? 0),
            fieldName: node?._$fieldName ?? ''
        });
        if (!sheet.SN._slicerCaches.has(cache.cacheId)) {
            sheet.SN._slicerCaches.set(cache.cacheId, cache);
        }

        const selectionNodes = ensureArray(node?.['sn:selection']?.['sn:item']);
        const selectedKeys = selectionNodes.map(sel => sel?._$key).filter(Boolean);

        const anchor = node?.['sn:anchor'];
        const startRow = Number(anchor?._$startRow ?? 0);
        const startCol = Number(anchor?._$startCol ?? 0);
        const offsetX = Number(anchor?._$offsetX ?? 0);
        const offsetY = Number(anchor?._$offsetY ?? 0);
        const width = Number(anchor?._$width ?? 180);
        const height = Number(anchor?._$height ?? 220);

        const drawing = sheet.Drawing._add({
            type: 'slicer',
            startCell: { r: startRow, c: startCol },
            offsetX,
            offsetY,
            width,
            height,
            // Keep slicer size when rows/cols are hidden, aligned with oneCell behavior.
            anchorType: 'oneCell'
        });

        const slicer = new SlicerItem(sheet, {
            id: node?._$id ?? `Slicer_${Date.now()}`,
            name: node?._$name ?? node?._$id,
            caption: node?._$caption ?? '',
            cacheId: cache.cacheId,
            sourceType: node?._$sourceType ?? 'table',
            sourceSheet: node?._$sheet ?? sheet.name,
            sourceName: node?._$sourceName ?? '',
            fieldIndex: Number(node?._$fieldIndex ?? 0),
            fieldName: node?._$fieldName ?? '',
            columns: Number(node?._$columns ?? 1),
            buttonHeight: Number(node?._$buttonHeight ?? 22),
            showHeader: node?._$showHeader !== '0',
            multiSelect: node?._$multiSelect === '1',
            styleName: node?._$styleName ?? 'SlicerStyleLight1',
            selectedKeys,
            drawing
        });

        sheet.Slicer._addParsed(slicer);
        if (selectedKeys.length > 0) {
            slicer.applyFilter();
        }
        parsedCount += 1;
    });

    return parsedCount;
}

function parseStandardSheetSlicers(sheet, _Xml) {
    const cacheMetaMap = sheet.SN._excelSlicerCacheMeta;
    if (!cacheMetaMap || cacheMetaMap.size === 0) return;

    const wsExtNodes = ensureArray(sheet._xmlObj?.extLst?.ext);
    const slicerRidList = [];

    wsExtNodes.forEach(ext => {
        const slicerList = ext?.['x14:slicerList'];
        if (!slicerList) return;
        ensureArray(slicerList?.['x14:slicer']).forEach(node => {
            const rid = node?.['_$r:id'];
            if (rid) slicerRidList.push(rid);
        });
    });

    if (slicerRidList.length === 0) return;

    const anchorMap = parseStandardSlicerAnchors(sheet, _Xml);

    slicerRidList.forEach((rid, ridIndex) => {
        const slicerPath = sheet.SN.Xml.getWsRelsFileTarget(sheet.rId, rid);
        const slicerRoot = slicerPath ? _Xml.obj[slicerPath]?.slicers : null;

        ensureArray(slicerRoot?.slicer).forEach((slicerNode, nodeIndex) => {
            const cacheMeta = cacheMetaMap.get(slicerNode?._$cache);
            const cache = ensureStandardCache(sheet, cacheMeta, slicerNode);

            const anchor = anchorMap.get(slicerNode?._$name) || {
                startRow: 0,
                startCol: 0,
                offsetX: 0,
                offsetY: 0,
                width: 180,
                height: 220
            };

            const drawing = sheet.Drawing._add({
                type: 'slicer',
                startCell: { r: anchor.startRow, c: anchor.startCol },
                offsetX: anchor.offsetX,
                offsetY: anchor.offsetY,
                width: anchor.width || 180,
                height: anchor.height || 220,
                anchorType: 'oneCell'
            });

            const buttonHeightPx = Math.max(
                18,
                Math.round(sheet.Utils.emuToPx(toNumber(slicerNode?._$rowHeight, 217488)))
            );

            const idBase = slicerNode?._$name || `Slicer_${ridIndex}_${nodeIndex}`;
            const slicer = new SlicerItem(sheet, {
                id: `Slicer_${idBase}_${ridIndex}_${nodeIndex}`,
                name: slicerNode?._$name || idBase,
                caption: slicerNode?._$caption || slicerNode?._$name || cache.fieldName,
                cacheId: cache.cacheId,
                sourceType: cache.sourceType,
                sourceSheet: cache.sourceSheet,
                sourceName: cache.sourceName,
                sourceId: cache.sourceId,
                fieldIndex: cache.fieldIndex,
                fieldName: cache.fieldName,
                columns: toNumber(slicerNode?._$columnCount, 1),
                buttonHeight: buttonHeightPx,
                showHeader: slicerNode?._$showCaption !== '0',
                multiSelect: false,
                styleName: slicerNode?._$style || 'SlicerStyleLight1',
                drawing
            });

            sheet.Slicer._addParsed(slicer);
        });
    });
}

export function parseSlicerCaches(SN, _Xml) {
    const root = _Xml.obj['xl/snSlicers.xml']?.['sn:slicers'];

    if (!SN._slicerCaches) SN._slicerCaches = new Map();
    SN._slicerCaches.clear();
    SN._slicerXmlNodes = [];

    if (root) {
        const cacheNodes = ensureArray(root['sn:caches']?.['sn:cache']);
        cacheNodes.forEach(node => {
            const cache = new SlicerCache(SN, {
                cacheId: Number(node?._$id),
                sourceType: node?._$type,
                sourceSheet: node?._$sheet,
                sourceName: node?._$name,
                fieldIndex: Number(node?._$fieldIndex ?? 0),
                fieldName: node?._$fieldName ?? ''
            });

            const items = parseCacheItems(node);
            cache.items = items;
            cache._itemMap = new Map(items.map(item => [item.key, item]));
            cache._dirty = false;
            cache._version = 1;
            SN._slicerCaches.set(cache.cacheId, cache);
        });

        SN._slicerXmlNodes = ensureArray(root['sn:slicer']);
    }

    parseStandardSlicerCachesMeta(SN, _Xml);
}

export function parseSheetSlicers(sheet, _Xml) {
    const customCount = parseCustomSheetSlicers(sheet);
    if (customCount > 0) return;
    parseStandardSheetSlicers(sheet, _Xml);
}
