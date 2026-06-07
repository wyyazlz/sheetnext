import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const XML_TEXT_RE = /\.(xml|rels|vml)$/i;
const TEXT_ENTRY_RE = /(^|\/)(\[Content_Types\]\.xml|[^/]+\.xml|[^/]+\.rels|[^/]+\.vml)$/i;
const WORKSHEET_ENTRY_RE = /^xl\/worksheets\/sheet\d+\.xml$/i;
const SHEET_DATA_RE = /<sheetData\b[^>]*>[\s\S]*?<\/sheetData>|<sheetData\b[^>]*\/>/i;
const SHEET_DATA_ROW_PARSER = new XMLParser({
    attributeNamePrefix: "_$",
    ignoreAttributes: false,
    parseAttributeValue: false,
    trimValues: false,
    parseTagValue: false,
    isArray: (tagName, jPath) => jPath === 'worksheet.sheetData.row'
});

function _postSuccess(payload, transfer = []) {
    self.postMessage({ type: 'success', payload }, transfer);
}

function _postProgress(stage, progress, extra = {}) {
    self.postMessage({
        type: 'progress',
        payload: { stage, progress, ...extra }
    });
}

function _postError(error) {
    self.postMessage({
        type: 'error',
        error: error?.message || String(error)
    });
}

function _normalizeBuffer(data) {
    if (data instanceof ArrayBuffer) return data;
    if (ArrayBuffer.isView(data)) {
        return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
    }
    return null;
}

function _copyArrayBuffer(data) {
    const buffer = _normalizeBuffer(data);
    if (!buffer) return null;
    return buffer.slice(0);
}

function _isTextEntry(name) {
    return XML_TEXT_RE.test(name) || TEXT_ENTRY_RE.test(name);
}

function _isWorksheetEntry(name) {
    return WORKSHEET_ENTRY_RE.test(name);
}

function _toArray(value) {
    if (!value) return [];
    return Array.isArray(value) ? value : [value];
}

function _extractSheetDataXml(text) {
    if (typeof text !== 'string') return '';
    return text.match(SHEET_DATA_RE)?.[0] || '';
}

function _stripSheetDataXml(text) {
    if (typeof text !== 'string') return '';
    return text.replace(SHEET_DATA_RE, '<sheetData/>');
}

function _parseSheetRows(sheetDataXml) {
    if (!sheetDataXml) return [];
    const parsed = SHEET_DATA_ROW_PARSER.parse(`<worksheet>${sheetDataXml}</worksheet>`);
    return _toArray(parsed?.worksheet?.sheetData?.row);
}

function _charToNum(value) {
    let num = 0;
    const text = String(value || '').toUpperCase();
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        if (code < 65 || code > 90) continue;
        num = num * 26 + (code - 64);
    }
    return Math.max(0, num - 1);
}

function _getCellColIndex(ref) {
    const match = String(ref || '').match(/[A-Z]+/i);
    return match ? _charToNum(match[0]) : 0;
}

function _getFormulaText(formulaNode) {
    if (!formulaNode) return '';
    if (typeof formulaNode === 'string') return formulaNode;
    return formulaNode['#text'] || '';
}

function _createSheetRowsMeta(rows) {
    const customRows = [];
    const formulas = [];
    const sharedFormulas = [];
    let maxRow = -1;
    let maxCol = -1;

    rows.forEach((row, fallbackIndex) => {
        const rIndex = Math.max(0, (Number.parseInt(row?.['_$r'], 10) || (fallbackIndex + 1)) - 1);
        maxRow = Math.max(maxRow, rIndex);

        if (
            row?.['_$ht'] !== undefined ||
            row?.['_$hidden'] === '1' ||
            row?.['_$outlineLevel'] !== undefined ||
            row?.['_$collapsed'] === '1' ||
            (row?.['_$s'] !== undefined && row?.['_$customFormat'] === '1')
        ) {
            customRows.push({
                r: rIndex,
                ht: row?.['_$ht'],
                hidden: row?.['_$hidden'] === '1',
                outlineLevel: row?.['_$outlineLevel'],
                collapsed: row?.['_$collapsed'] === '1',
                styleIndex: row?.['_$s'],
                customFormat: row?.['_$customFormat'] === '1'
            });
        }

        _toArray(row?.c).forEach(cell => {
            const cIndex = _getCellColIndex(cell?.['_$r']);
            maxCol = Math.max(maxCol, cIndex);
            const formulaNode = cell?.f;
            if (!formulaNode) return;

            const formulaText = _getFormulaText(formulaNode);
            const sharedIndex = typeof formulaNode === 'object' ? formulaNode['_$si'] : undefined;
            const formulaInfo = { r: rIndex, c: cIndex, f: formulaText };
            if (sharedIndex !== undefined) formulaInfo.si = sharedIndex;
            if (formulaText || sharedIndex !== undefined) formulas.push(formulaInfo);

            if (typeof formulaNode === 'object' && formulaNode['_$t'] === 'shared' && sharedIndex !== undefined && formulaText) {
                sharedFormulas.push(formulaInfo);
            }
        });
    });

    return {
        rowCount: maxRow + 1,
        colCount: maxCol + 1,
        customRows,
        formulas,
        sharedFormulas
    };
}

function _importXlsx(payload) {
    const buffer = _normalizeBuffer(payload?.buffer);
    if (!buffer) throw new Error('xlsx_worker_empty_import_buffer');

    _postProgress('openingPackage', 14);
    const zip = new JSZip();
    zip.load(buffer);

    const keys = Object.keys(zip.files || {}).filter(key => !zip.files[key]?.dir);
    const entries = [];
    const sheetRows = Object.create(null);
    const sheetRowMeta = Object.create(null);
    const transfer = [];
    const total = Math.max(1, keys.length);

    keys.forEach((key, index) => {
        const item = zip.files[key];
        if (!item) return;

        if (_isTextEntry(key)) {
            const text = item.asText();
            if (_isWorksheetEntry(key)) {
                const sheetDataXml = _extractSheetDataXml(text);
                const rows = _parseSheetRows(sheetDataXml);
                if (rows.length > 0) {
                    sheetRows[key] = rows;
                    sheetRowMeta[key] = _createSheetRowsMeta(rows);
                }
                entries.push({ name: key, type: 'text', data: _stripSheetDataXml(text) });
            } else {
                entries.push({ name: key, type: 'text', data: text });
            }
        } else {
            const entryBuffer = item.asArrayBuffer();
            const copied = _copyArrayBuffer(entryBuffer);
            if (copied) {
                entries.push({ name: key, type: 'binary', data: copied });
                transfer.push(copied);
            }
        }

        if (index % 20 === 0 || index === keys.length - 1) {
            _postProgress('parsing', 14 + ((index + 1) / total) * 70, {
                current: index + 1,
                total
            });
        }
    });

    _postSuccess({ entries, sheetRows, sheetRowMeta }, transfer);
}

function _zipEntries(payload) {
    const entries = Array.isArray(payload?.entries) ? payload.entries : [];
    if (entries.length === 0) throw new Error('xlsx_worker_empty_export_entries');

    const zip = new JSZip();
    const total = Math.max(1, entries.length);

    entries.forEach((entry, index) => {
        if (!entry?.name) return;
        if (entry.type === 'binary') {
            const buffer = _normalizeBuffer(entry.data);
            if (buffer) zip.file(entry.name, new Uint8Array(buffer), { binary: true });
        } else {
            zip.file(entry.name, entry.data ?? '');
        }

        if (index % 20 === 0 || index === entries.length - 1) {
            _postProgress('preparingPackage', 76 + ((index + 1) / total) * 8, {
                current: index + 1,
                total
            });
        }
    });

    _postProgress('compressing', 88);
    const uint8 = zip.generate({ type: 'uint8array', compression: 'DEFLATE' });
    const buffer = uint8.buffer.slice(uint8.byteOffset, uint8.byteOffset + uint8.byteLength);
    _postSuccess({ buffer }, [buffer]);
}

self.onmessage = (event) => {
    const { type, payload } = event.data || {};
    try {
        if (type === 'import-xlsx') {
            _importXlsx(payload);
        } else if (type === 'zip-entries') {
            _zipEntries(payload);
        } else {
            throw new Error(`unknown_xlsx_worker_task:${type}`);
        }
    } catch (error) {
        _postError(error);
    }
};
