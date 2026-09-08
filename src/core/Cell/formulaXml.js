import { FORMULA_DATA } from '../../assets/formulaData.js';
import { parseFormula, tokenizeFormula } from '../Formula/FormulaAst.js';

const FUTURE_FUNCTIONS = new Set([
    ...FORMULA_DATA.filter(item => Number(item.version) >= 2010 || item.version === 'Microsoft 365').map(item => item.name),
    'AGGREGATE', 'BASE', 'CEILING.PRECISE', 'FLOOR.PRECISE', 'ISO.CEILING'
]);
const KNOWN_FUNCTIONS = new Set([...FORMULA_DATA.map(item => item.name), ...FUTURE_FUNCTIONS]);
const WORKSHEET_FUNCTIONS = new Set(['FILTER', 'SORT']);
const ARRAY_FUNCTIONS = new Set([
    'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY',
    'CHOOSECOLS', 'CHOOSEROWS', 'DROP', 'EXPAND', 'HSTACK', 'VSTACK',
    'TAKE', 'TEXTSPLIT', 'TOCOL', 'TOROW', 'WRAPCOLS', 'WRAPROWS',
    'BYCOL', 'BYROW', 'MAKEARRAY', 'MAP', 'SCAN', 'GROUPBY', 'PIVOTBY', 'TRIMRANGE'
]);
const METADATA_PART = 'xl/metadata.xml';
const METADATA_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata';

function _bareFunctionName(name) {
    return name.replace(/^(?:(?:_xlfn|_xlws|_xludf)\.)+/i, '').toUpperCase();
}

function _mapFunctionNames(formula, exporting) {
    const tokens = tokenizeFormula(formula);
    const pieces = [];
    let offset = 0;
    tokens.forEach((token, i) => {
        if (token.type !== 'IDENT' || tokens[i + 1]?.type !== 'LPAREN') return;
        const name = _bareFunctionName(token.value);
        if (!KNOWN_FUNCTIONS.has(name)) return;
        const prefix = exporting && FUTURE_FUNCTIONS.has(name)
            ? (WORKSHEET_FUNCTIONS.has(name) ? '_xlfn._xlws.' : '_xlfn.')
            : '';
        pieces.push(formula.slice(offset, token.start), prefix + name);
        offset = token.end;
    });
    pieces.push(formula.slice(offset));
    return pieces.join('');
}

export function _normalizeFormulaFromXlsx(formula) {
    return _mapFunctionNames(String(formula ?? ''), false);
}

function _isArrayFormula(formula) {
    try {
        const ast = parseFormula(formula);
        return ast.type === 'CallExpression' && ARRAY_FUNCTIONS.has(_bareFunctionName(ast.name));
    } catch {
        return false;
    }
}

export function _writeFormulaCachedValue(obj, cell) {
    const value = cell.calcVal;
    if (cell.type === 'error' || value instanceof Error) {
        obj._$t = 'e';
        obj.v = value instanceof Error ? value.message : String(value);
    } else if (typeof value === 'boolean') {
        obj._$t = 'b';
        obj.v = value ? '1' : '0';
    } else if (typeof value === 'number' && Number.isFinite(value)) {
        obj.v = value;
    } else if (typeof value === 'string') {
        obj._$t = 'str';
        obj.v = value;
    } else {
        obj.v = '';
    }
}

export function _buildFormulaXml(obj, cell, rowIndex, colIndex) {
    _writeFormulaCachedValue(obj, cell);
    const formula = cell.editVal.substring(1);
    const text = _mapFunctionNames(formula, true);
    const imported = cell._xmlObj?.f;
    const importedArray = imported?._$t === 'array'
        && _normalizeFormulaFromXlsx(imported['#text']) === _normalizeFormulaFromXlsx(formula);
    const spill = cell._spillRange;
    const dynamic = !!spill || _isArrayFormula(formula) || (importedArray && Number(cell._xmlObj?._$cm) > 0);
    if (dynamic || importedArray) {
        const end = spill
            ? `${cell._SN.Utils.numToChar(colIndex + spill.cols - 1)}${rowIndex + spill.rows}`
            : null;
        const range = end ? `${obj._$r}:${end}` : (importedArray ? imported._$ref : obj._$r);
        obj.f = { '_$t': 'array', '_$ref': range || obj._$r, '#text': text };
        if (dynamic) {
            obj._$cm = '1';
            cell._SN._Xml._hasDynamicArrayFormulas = true;
        }
    } else {
        obj.f = text;
    }
}

export function _prepareFormulaExport(SN) {
    // Spill evaluation can create rows, so finish it before worksheet serialization starts.
    const formulas = [];
    SN.sheets.forEach(sheet => sheet.rows.forEach(row => row?.cells?.forEach(cell => {
        if (cell?.isFormula) formulas.push(cell);
    })));
    formulas.forEach(cell => { void cell.calcVal; });
}

export function _exportFormulaMetadata(xml) {
    if (!xml._hasDynamicArrayFormulas) return;
    xml.obj[METADATA_PART] = {
        '?xml': { '_$version': '1.0', '_$encoding': 'UTF-8', '_$standalone': 'yes' },
        metadata: {
            '_$xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            '_$xmlns:xlrd': 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray',
            metadataTypes: {
                '_$count': '1',
                metadataType: {
                    '_$name': 'XLDAPR', '_$minSupportedVersion': '120000',
                    '_$copy': '1', '_$pasteAll': '1', '_$pasteValues': '1', '_$merge': '1',
                    '_$splitFirst': '1', '_$rowColShift': '1', '_$clearFormats': '1',
                    '_$clearComments': '1', '_$assign': '1', '_$coerce': '1', '_$cellMeta': '1'
                }
            },
            futureMetadata: {
                '_$name': 'XLDAPR', '_$count': '1',
                bk: { extLst: { ext: {
                    '_$uri': '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}',
                    'xlrd:dynamicArrayProperties': { '_$fDynamic': '1', '_$fCollapsed': '0' }
                } } }
            },
            cellMetadata: { '_$count': '1', bk: { rc: { '_$t': '1', '_$v': '0' } } }
        }
    };
    const relationships = xml.relationship;
    if (!relationships.some(rel => rel._$Type === METADATA_REL)) {
        const nextId = relationships.reduce((max, rel) => {
            const match = /^rId(\d+)$/.exec(rel._$Id || '');
            return match ? Math.max(max, Number(match[1])) : max;
        }, 0) + 1;
        relationships.push({ '_$Id': `rId${nextId}`, '_$Type': METADATA_REL, '_$Target': 'metadata.xml' });
    }
    const overrides = xml.obj['[Content_Types].xml'].Types.Override;
    if (!overrides.some(item => item._$PartName === `/${METADATA_PART}`)) {
        overrides.push({
            '_$PartName': `/${METADATA_PART}`,
            '_$ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml'
        });
    }
}
