import { toRef } from './TableUtils.js';

function parseOptionalInt(value) {
    if (value === undefined || value === null || value === '') return undefined;
    const parsed = Number.parseInt(value, 10);
    return Number.isNaN(parsed) ? undefined : parsed;
}

function resolveExportDxfId(dxfId, mapDxfId) {
    if (!Number.isInteger(dxfId)) return undefined;
    if (typeof mapDxfId !== 'function') return dxfId;

    const mappedId = mapDxfId(dxfId);
    return Number.isInteger(mappedId) ? mappedId : undefined;
}

/**
 * Build table xml object.
 * @param {TableItem} table
 * @param {number} tableIndex
 * @param {Object} [options]
 * @param {(dxfId:number)=>number|undefined} [options.mapDxfId]
 * @returns {Object|null}
 */
export function buildTableXml(table, tableIndex, options = {}) {
    const range = table.range;
    if (!range) return null;
    const mapDxfId = options.mapDxfId;
    const buildDxfAttr = (name, dxfId) => {
        const exportDxfId = resolveExportDxfId(dxfId, mapDxfId);
        return Number.isInteger(exportDxfId) ? { [name]: exportDxfId } : {};
    };

    // Keep column names synced with current header cells before export.
    table.syncColumnNames();

    const ref = table.ref;
    const id = tableIndex + 1;

    const tableColumns = table.columns.map((col, i) => ({
        '_$id': col.id || (i + 1),
        '_$name': col.name || `Column${i + 1}`,
        ...buildDxfAttr('_$dataDxfId', col.dataDxfId),
        ...(col.totalsRowFunction ? { '_$totalsRowFunction': col.totalsRowFunction } : {}),
        ...(col.totalsRowLabel ? { '_$totalsRowLabel': col.totalsRowLabel } : {})
    }));

    const filterRef = table.showTotalsRow
        ? toRef({ s: range.s, e: { r: range.e.r - 1, c: range.e.c } })
        : ref;

    const tableScopeId = table.sheet?.AutoFilter?.getTableScopeId?.(table.id);
    const scopeFilterXml = tableScopeId ? table.sheet.AutoFilter.buildXml(tableScopeId) : null;
    const tableAutoFilterXml = scopeFilterXml
        ? { ...scopeFilterXml, '_$ref': filterRef }
        : { '_$ref': filterRef };

    const tableObj = {
        '_$xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        '_$id': id,
        '_$name': table.name,
        '_$displayName': table.displayName || table.name,
        '_$ref': ref,
        '_$totalsRowShown': '0',
        ...buildDxfAttr('_$headerRowDxfId', table.headerRowDxfId),
        ...buildDxfAttr('_$dataDxfId', table.dataDxfId),
        ...buildDxfAttr('_$headerRowBorderDxfId', table.headerRowBorderDxfId),
        ...buildDxfAttr('_$tableBorderDxfId', table.tableBorderDxfId),
        ...buildDxfAttr('_$totalsRowBorderDxfId', table.totalsRowBorderDxfId),
        ...(table.showHeaderRow ? {} : { '_$headerRowCount': '0' }),
        ...(table.showTotalsRow ? { '_$totalsRowCount': '1' } : {}),
        ...(table.autoFilterEnabled ? { autoFilter: tableAutoFilterXml } : {}),
        tableColumns: {
            '_$count': tableColumns.length,
            tableColumn: tableColumns
        },
        tableStyleInfo: {
            '_$name': table.styleId,
            '_$showFirstColumn': table.showFirstColumn ? '1' : '0',
            '_$showLastColumn': table.showLastColumn ? '1' : '0',
            '_$showRowStripes': table.showRowStripes ? '1' : '0',
            '_$showColumnStripes': table.showColumnStripes ? '1' : '0'
        }
    };

    return {
        '?xml': { '_$version': '1.0', '_$encoding': 'UTF-8', '_$standalone': 'yes' },
        table: tableObj
    };
}

/**
 * Export all tables from a sheet.
 * @param {Sheet} sheet
 * @param {Object} _Xml
 * @param {Array} sheetRelationship
 * @param {number} sheetIndex
 * @param {Object} [options]
 * @param {(dxfId:number)=>number|undefined} [options.mapDxfId]
 * @returns {{count:number, tableParts:Array}|number}
 */
export function exportTables(sheet, _Xml, sheetRelationship, sheetIndex, options = {}) {
    const tables = sheet.Table;
    if (!tables || tables.size === 0) return 0;
    const mapDxfId = options.mapDxfId;

    let tableCount = 0;
    const tableParts = [];

    tables.forEach(table => {
        tableCount += 1;
        const tableIndex = sheetIndex * 100 + tableCount;
        const tableFileName = `table${tableIndex}.xml`;
        const tableXml = buildTableXml(table, tableIndex, { mapDxfId });
        if (!tableXml) return;

        _Xml.obj[`xl/tables/${tableFileName}`] = tableXml;

        _Xml.obj['[Content_Types].xml'].Types.Override.push({
            '_$PartName': `/xl/tables/${tableFileName}`,
            '_$ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
        });

        const rId = `rId${sheetRelationship.length + 1}`;
        sheetRelationship.push({
            '_$Id': rId,
            '_$Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
            '_$Target': `../tables/${tableFileName}`
        });

        tableParts.push({ '_$r:id': rId });
    });

    return { count: tableCount, tableParts };
}

/**
 * Parse one table xml.
 * @param {Sheet} sheet
 * @param {Object} tableXml
 * @returns {TableItem|null}
 */
export function parseTableXml(sheet, tableXml) {
    const tableNode = tableXml?.table;
    if (!tableNode) return null;

    const ref = tableNode['_$ref'];
    const name = tableNode['_$name'];
    const displayName = tableNode['_$displayName'] || name;
    const id = tableNode['_$id'];

    const columnsNode = tableNode.tableColumns?.tableColumn;
    const columnsList = Array.isArray(columnsNode) ? columnsNode : (columnsNode ? [columnsNode] : []);
    const columns = columnsList.map(col => ({
        id: Number.parseInt(col['_$id'], 10) || 0,
        name: col['_$name'] || '',
        dataDxfId: parseOptionalInt(col['_$dataDxfId']),
        totalsRowFunction: col['_$totalsRowFunction'] || null,
        totalsRowLabel: col['_$totalsRowLabel'] || null
    }));

    const styleInfo = tableNode.tableStyleInfo || {};
    const styleId = styleInfo['_$name'] || 'TableStyleMedium2';
    const showFirstColumn = styleInfo['_$showFirstColumn'] === '1';
    const showLastColumn = styleInfo['_$showLastColumn'] === '1';
    const showRowStripes = styleInfo['_$showRowStripes'] !== '0';
    const showColumnStripes = styleInfo['_$showColumnStripes'] === '1';

    const headerRowCount = Number.parseInt(tableNode['_$headerRowCount'], 10) || 1;
    const totalsRowCount = Number.parseInt(tableNode['_$totalsRowCount'], 10) || 0;
    const showHeaderRow = headerRowCount > 0;
    const showTotalsRow = totalsRowCount > 0;

    return sheet.Table.add({
        id: `table_${id}`,
        name,
        displayName,
        rangeRef: ref,
        columns,
        styleId,
        showHeaderRow,
        showTotalsRow,
        showFirstColumn,
        showLastColumn,
        showRowStripes,
        showColumnStripes,
        autoFilterEnabled: !!tableNode.autoFilter,
        headerRowDxfId: parseOptionalInt(tableNode['_$headerRowDxfId']),
        dataDxfId: parseOptionalInt(tableNode['_$dataDxfId']),
        headerRowBorderDxfId: parseOptionalInt(tableNode['_$headerRowBorderDxfId']),
        tableBorderDxfId: parseOptionalInt(tableNode['_$tableBorderDxfId']),
        totalsRowBorderDxfId: parseOptionalInt(tableNode['_$totalsRowBorderDxfId']),
        _tableAutoFilterXml: tableNode.autoFilter || null,
        _fromImport: true
    });
}

/**
 * Parse all table parts linked from one sheet.
 * @param {Sheet} sheet
 * @param {Object} xmlObj
 * @param {string} sheetFileName
 */
export function parseSheetTables(sheet, xmlObj, sheetFileName) {
    const relsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;
    const rels = xmlObj[relsPath]?.Relationships?.Relationship;
    if (!rels) return;

    const relsList = Array.isArray(rels) ? rels : [rels];
    relsList.forEach(rel => {
        if (!rel['_$Type']?.includes('/table')) return;

        let target = rel['_$Target'];
        if (target.startsWith('../')) {
            target = `xl/${target.substring(3)}`;
        } else if (!target.startsWith('xl/')) {
            target = `xl/tables/${target}`;
        }

        const tableXml = xmlObj[target];
        if (tableXml) parseTableXml(sheet, tableXml);
    });
}

