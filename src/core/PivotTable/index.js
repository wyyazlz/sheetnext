import PivotTable from './PivotTable.js';
import PivotDataField from './PivotDataField.js';

export { PivotTable };
export { PivotDataField };

export class PivotTableItem {}
export class PivotCache {}
export class PivotField {}

export function initPivotTableModule(SN) {
    if (!SN._pivotCaches) SN._pivotCaches = new Map();
}
