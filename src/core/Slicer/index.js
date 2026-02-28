import SlicerItem from './SlicerItem.js';
import SlicerCache from './SlicerCache.js';
import Slicer from './Slicer.js';

const DEFAULT_SLICER_COL_WIDTH = 2.6;
const DEFAULT_SLICER_ROW_HEIGHT = 11;

function getDefaultSlicerSize(sheet) {
    const colWidth = sheet?.defaultColWidth || 70;
    const rowHeight = sheet?.defaultRowHeight || 20;
    return {
        width: Math.round(colWidth * DEFAULT_SLICER_COL_WIDTH),
        height: Math.round(rowHeight * DEFAULT_SLICER_ROW_HEIGHT)
    };
}

function calculateRowStep(sheet, height) {
    const baseHeight = sheet.defaultRowHeight || 20;
    return Math.max(1, Math.ceil((height + 10) / baseHeight));
}

export function createTableSlicers(SN, sheet, table, columns = []) {
    const baseStart = {
        r: table.range?.s?.r ?? sheet.activeCell.r,
        c: (table.range?.e?.c ?? sheet.activeCell.c) + 2
    };

    const { width: slicerWidth, height: slicerHeight } = getDefaultSlicerSize(sheet);
    const rowStep = calculateRowStep(sheet, slicerHeight);

    const slicers = [];
    columns.forEach((col, index) => {
        const startCell = { r: baseStart.r + index * rowStep, c: baseStart.c };
        const slicer = sheet.Slicer.add({
            sourceType: 'table',
            sourceSheet: sheet.name,
            sourceName: table.name,
            sourceId: table.id,
            fieldIndex: col.index,
            fieldName: col.name,
            caption: col.name,
            startCell,
            width: slicerWidth,
            height: slicerHeight
        });
        if (slicer) slicers.push(slicer);
    });

    if (SN.Canvas?.r) SN.Canvas.r();
    return slicers;
}

export function createPivotSlicers(SN, sheet, pivotTable, fields = []) {
    const locationRef = pivotTable.location?.ref || 'A1';
    const range = sheet.rangeStrToNum(locationRef);
    const baseStart = {
        r: range?.s?.r ?? sheet.activeCell.r,
        c: (range?.e?.c ?? sheet.activeCell.c) + 2
    };

    const { width: slicerWidth, height: slicerHeight } = getDefaultSlicerSize(sheet);
    const rowStep = calculateRowStep(sheet, slicerHeight);
    const slicers = [];

    fields.forEach((fieldIndex, index) => {
        const field = pivotTable.pivotFields?.[fieldIndex];
        const fieldName = field?.name ?? `Field ${fieldIndex + 1}`;
        const startCell = { r: baseStart.r + index * rowStep, c: baseStart.c };

        const slicer = sheet.Slicer.add({
            sourceType: 'pivot',
            sourceSheet: sheet.name,
            sourceName: pivotTable.name,
            fieldIndex,
            fieldName,
            caption: fieldName,
            startCell,
            width: slicerWidth,
            height: slicerHeight
        });
        if (slicer) slicers.push(slicer);
    });

    if (SN.Canvas?.r) SN.Canvas.r();
    return slicers;
}

export { SlicerItem, SlicerCache, Slicer };
