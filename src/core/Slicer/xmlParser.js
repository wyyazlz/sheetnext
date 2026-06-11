import { buildSlicerKey, parseSlicerKey, formatSlicerValue } from './helpers.js';
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

function parseCustomPivotTargets(node, fallback = {}) {
    const targets = [];
    const pivotNodes = ensureArray(node?.['sn:pivotTables']?.['sn:pivotTable']);
    pivotNodes.forEach(item => {
        const sourceSheet = item?._$sheet || fallback.sourceSheet || '';
        const sourceName = item?._$name || fallback.sourceName || '';
        if (!sourceSheet || !sourceName) return;
        targets.push({
            sourceSheet,
            sourceName,
            sourceId: sourceName,
            fieldIndex: toNumber(item?._$fieldIndex, fallback.fieldIndex ?? 0)
        });
    });
    return targets;
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

function findPivotTableByName(SN, name) {
    if (!name) return null;
    for (const sheet of (SN.sheets || [])) {
        const pivotTable = sheet.PivotTable?.get?.(name);
        if (pivotTable) return pivotTable;
    }
    return null;
}

function getAllPivotTables(SN) {
    const result = [];
    (SN.sheets || []).forEach(sheet => {
        sheet.PivotTable?.getAll?.().forEach(pivotTable => result.push(pivotTable));
    });
    return result;
}

function resolvePivotTableFromCache(SN, cacheDef) {
    const pivotRefs = ensureArray(cacheDef?.pivotTables?.pivotTable);
    for (const pivotRef of pivotRefs) {
        const pivotTable = findPivotTableByName(SN, pivotRef?._$name);
        if (pivotTable) return pivotTable;
    }

    const pivotTables = getAllPivotTables(SN);
    return pivotTables.length === 1 ? pivotTables[0] : null;
}

function resolvePivotTablesFromCache(SN, cacheDef) {
    const pivotRefs = ensureArray(cacheDef?.pivotTables?.pivotTable);
    const result = [];

    pivotRefs.forEach(pivotRef => {
        const pivotTable = findPivotTableByName(SN, pivotRef?._$name);
        if (pivotTable && !result.includes(pivotTable)) result.push(pivotTable);
    });

    if (result.length > 0) return result;

    const pivotTables = getAllPivotTables(SN);
    return pivotTables.length === 1 ? pivotTables : [];
}

function resolvePivotFieldIndex(pivotTable, sourceName) {
    const name = String(sourceName || '').trim();
    if (!pivotTable || !name) return 0;

    const cacheFields = pivotTable.cache?.fields || [];
    const cacheIndex = cacheFields.findIndex(field => field?.name === name);
    if (cacheIndex >= 0) return cacheIndex;

    const pivotIndex = pivotTable.pivotFields?.findIndex?.(field => field?.name === name);
    return pivotIndex >= 0 ? pivotIndex : 0;
}

function buildPivotSlicerCacheItems(pivotTable, fieldIndex, cacheDef) {
    const sharedItems = pivotTable?.cache?.fields?.[fieldIndex]?.sharedItems?.values || [];
    const itemNodes = ensureArray(cacheDef?.data?.tabular?.items?.i);
    const sourceNodes = itemNodes.length > 0
        ? itemNodes
        : sharedItems.map((_, index) => ({ _$x: String(index) }));

    const items = [];
    const itemMap = new Map();
    const selectedKeys = [];
    let hasSelectedState = false;

    sourceNodes.forEach(node => {
        const sharedIndex = toNumber(node?._$x, NaN);
        if (!Number.isFinite(sharedIndex)) return;

        const sharedItem = sharedItems[sharedIndex];
        const value = sharedItem?.value ?? null;
        const key = buildSlicerKey(value);
        if (itemMap.has(key)) return;

        const item = {
            key,
            value,
            text: pivotTable?._getItemDisplayValue?.(fieldIndex, sharedIndex) ?? formatSlicerValue(value),
            sharedIndex,
            count: 0
        };

        itemMap.set(key, item);
        items.push(item);

        if (node?._$s === '1') {
            hasSelectedState = true;
            selectedKeys.push(key);
        }
    });

    return {
        items,
        itemMap,
        selectedKeys: hasSelectedState ? selectedKeys : []
    };
}

function applyParsedCacheItems(cache, cacheMeta) {
    if (!cache || !cacheMeta?.items) return;
    cache.items = cacheMeta.items;
    cache._itemMap = cacheMeta.itemMap || new Map(cache.items.map(item => [item.key, item]));
    cache._dirty = false;
    cache._version = 1;
}

function ensureStandardCache(sheet, cacheMeta, slicerNode) {
    const SN = sheet.SN;
    if (!SN._slicerCaches) SN._slicerCaches = new Map();

    const cacheId = cacheMeta?.cacheId;
    if (Number.isFinite(cacheId) && SN._slicerCaches.has(cacheId)) {
        const existing = SN._slicerCaches.get(cacheId);
        if (cacheMeta?.cacheName && !existing._excelCacheName) existing._excelCacheName = cacheMeta.cacheName;
        if (cacheMeta?.uid && !existing._excelUid) existing._excelUid = cacheMeta.uid;
        if (Array.isArray(cacheMeta?.selectedKeys) && existing.selectedKeys.size === 0) {
            existing.selectedKeys = cacheMeta.selectedKeys;
        }
        applyParsedCacheItems(existing, cacheMeta);
        return existing;
    }

    if (cacheMeta?.sourceType === 'pivot') {
        const cache = new SlicerCache(SN, {
            cacheId,
            sourceType: 'pivot',
            sourceSheet: cacheMeta.sourceSheet || sheet.name,
            sourceName: cacheMeta.sourceName || null,
            sourceId: cacheMeta.sourceId || null,
            fieldIndex: Number.isFinite(cacheMeta.fieldIndex) ? cacheMeta.fieldIndex : 0,
            fieldName: cacheMeta.fieldName || slicerNode?._$name || '',
            pivotTables: cacheMeta.pivotTables || [],
            selectedKeys: cacheMeta.selectedKeys || []
        });
        cache._excelCacheName = cacheMeta.cacheName || null;
        cache._excelUid = cacheMeta.uid || null;
        applyParsedCacheItems(cache, cacheMeta);
        SN._slicerCaches.set(cache.cacheId, cache);
        return cache;
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
        fieldName,
        selectedKeys: cacheMeta?.selectedKeys || []
    });
    cache._excelCacheName = cacheMeta?.cacheName || null;
    cache._excelUid = cacheMeta?.uid || null;
    applyParsedCacheItems(cache, cacheMeta);

    SN._slicerCaches.set(cache.cacheId, cache);
    return cache;
}

function parseStandardSlicerCachesMeta(SN, _Xml) {
    const workbook = _Xml.obj['xl/workbook.xml']?.workbook;
    const extNodes = ensureArray(workbook?.extLst?.ext);
    const cacheRidList = [];

    extNodes.forEach(ext => {
        const slicerCachesNode = ext?.['x14:slicerCaches'] || ext?.['x15:slicerCaches'];
        if (!slicerCachesNode) return;
        const cacheRefs = slicerCachesNode?.['x14:slicerCache'] || slicerCachesNode?.['x15:slicerCache'];
        ensureArray(cacheRefs).forEach(cacheRef => {
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
        const tableMeta = tableSlicerCache ? {
            sourceType: 'table',
            sourceName: cacheDef?._$sourceName || '',
            tableId: toNumber(tableSlicerCache?._$tableId, NaN),
            column: toNumber(tableSlicerCache?._$column, NaN)
        } : null;

        const pivotTables = tableMeta ? [] : resolvePivotTablesFromCache(SN, cacheDef);
        const pivotTable = pivotTables[0] || null;
        const fieldIndex = pivotTable ? resolvePivotFieldIndex(pivotTable, cacheDef?._$sourceName) : 0;
        const pivotTargets = pivotTables.map(item => ({
            sourceSheet: item.sheet?.name || '',
            sourceName: item.name,
            sourceId: item.name,
            fieldIndex: resolvePivotFieldIndex(item, cacheDef?._$sourceName)
        }));
        const pivotItems = pivotTable ? buildPivotSlicerCacheItems(pivotTable, fieldIndex, cacheDef) : null;
        const pivotMeta = pivotTable ? {
            sourceType: 'pivot',
            sourceSheet: pivotTable.sheet?.name || '',
            sourceName: pivotTable.name,
            sourceId: pivotTable.name,
            fieldIndex,
            fieldName: pivotTable.cache?.fields?.[fieldIndex]?.name || cacheDef?._$sourceName || '',
            pivotTables: pivotTargets,
            items: pivotItems?.items || [],
            itemMap: pivotItems?.itemMap || new Map(),
            selectedKeys: pivotItems?.selectedKeys || []
        } : {
            sourceType: 'pivot',
            sourceSheet: '',
            sourceName: '',
            sourceId: '',
            fieldIndex: 0,
            fieldName: cacheDef?._$sourceName || '',
            pivotTables: [],
            items: [],
            itemMap: new Map(),
            selectedKeys: []
        };

        SN._excelSlicerCacheMeta.set(cacheName, {
            cacheId: nextCacheId++,
            cacheName,
            uid: cacheDef?._$xr10_uid ?? cacheDef?._$xr10Uid ?? cacheDef?.['_$xr10:uid'] ?? null,
            ...(tableMeta || pivotMeta)
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
        const slicerList = ext?.['x14:slicerList'] || ext?.['x15:slicerList'];
        if (!slicerList) return;
        const slicerNodes = slicerList?.['x14:slicer'] || slicerList?.['x15:slicer'];
        ensureArray(slicerNodes).forEach(node => {
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
                selectedKeys: cacheMeta?.selectedKeys || [],
                drawing
            });

            sheet.Slicer._addParsed(slicer);
            if (slicer.selectedKeys.size > 0) {
                slicer.applyFilter();
            }
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
            const sourceSheet = node?._$sheet;
            const sourceName = node?._$name;
            const fieldIndex = Number(node?._$fieldIndex ?? 0);
            const cache = new SlicerCache(SN, {
                cacheId: Number(node?._$id),
                sourceType: node?._$type,
                sourceSheet,
                sourceName,
                fieldIndex,
                fieldName: node?._$fieldName ?? '',
                pivotTables: parseCustomPivotTargets(node, { sourceSheet, sourceName, fieldIndex })
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
