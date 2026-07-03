function _normalizeUsedRangeOptions(options = {}) {
    const normalized = options && typeof options === 'object' ? options : {};
    return {
        contents: normalized.contents !== false,
        formulas: normalized.formulas !== false,
        richText: normalized.richText !== false,
        hyperlinks: normalized.hyperlinks !== false,
        dataValidations: normalized.dataValidations !== false,
        controls: normalized.controls !== false,
        comments: normalized.comments !== false,
        merges: normalized.merges !== false,
        tables: normalized.tables !== false,
        autoFilters: normalized.autoFilters !== false,
        drawings: normalized.drawings === true,
        sparklines: normalized.sparklines === true,
        conditionalFormats: normalized.conditionalFormats === true,
        pivotTables: normalized.pivotTables === true,
        includeFormats: normalized.includeFormats === true,
        includeRowColFormats: normalized.includeRowColFormats === true,
        emptyAsNull: normalized.emptyAsNull === true,
        asString: normalized.asString === true
    };
}

function _createBounds() {
    return {
        minR: Infinity,
        minC: Infinity,
        maxR: -1,
        maxC: -1
    };
}

function _isValidIndex(value) {
    return Number.isFinite(value) && value >= 0;
}

function _addCell(bounds, r, c) {
    const row = Number(r);
    const col = Number(c);
    if (!_isValidIndex(row) || !_isValidIndex(col)) return;
    bounds.minR = Math.min(bounds.minR, Math.floor(row));
    bounds.minC = Math.min(bounds.minC, Math.floor(col));
    bounds.maxR = Math.max(bounds.maxR, Math.floor(row));
    bounds.maxC = Math.max(bounds.maxC, Math.floor(col));
}

function _normalizeRange(range) {
    if (!range?.s || !range?.e) return null;
    const sr = Number(range.s.r);
    const sc = Number(range.s.c);
    const er = Number(range.e.r);
    const ec = Number(range.e.c);
    if (!_isValidIndex(sr) || !_isValidIndex(sc) || !_isValidIndex(er) || !_isValidIndex(ec)) return null;
    return {
        s: { r: Math.floor(Math.min(sr, er)), c: Math.floor(Math.min(sc, ec)) },
        e: { r: Math.floor(Math.max(sr, er)), c: Math.floor(Math.max(sc, ec)) }
    };
}

function _addRange(bounds, range) {
    const normalized = _normalizeRange(range);
    if (!normalized) return;
    _addCell(bounds, normalized.s.r, normalized.s.c);
    _addCell(bounds, normalized.e.r, normalized.e.c);
}

function _stripSheetPrefix(ref) {
    const text = String(ref ?? '').trim();
    const bangIndex = text.lastIndexOf('!');
    return bangIndex >= 0 ? text.slice(bangIndex + 1) : text;
}

function _addRef(bounds, sheet, ref) {
    const clean = _stripSheetPrefix(ref);
    if (!clean) return;
    try {
        _addRange(bounds, sheet.rangeStrToNum(clean));
    } catch (_error) { }
}

function _addSqref(bounds, sheet, sqref) {
    String(sqref ?? '').split(/\s+/).forEach(ref => _addRef(bounds, sheet, ref));
}

function _cellHasContent(cell, options) {
    if (!cell) return false;
    const value = cell._editVal;
    const isFormula = typeof value === 'string' && value.startsWith('=');
    if (options.formulas && isFormula) return true;
    if (isFormula) return false;
    if (options.contents && value !== undefined && value !== null && value !== '') return true;
    if (options.richText && cell._richText) return true;
    if (options.hyperlinks && cell._hyperlink) return true;
    if (options.dataValidations && cell._dataValidation) return true;
    if (options.controls && cell._style?.control) return true;
    if (options.includeFormats && (
        (cell._style && Object.keys(cell._style).length > 0) ||
        cell._xmlObj?.['_$s'] !== undefined
    )) return true;
    return false;
}

function _addExistingCells(bounds, sheet, options) {
    sheet.rows.forEach((row, rIndex) => {
        if (!row?.cells) return;
        row.cells.forEach((cell, cIndex) => {
            if (_cellHasContent(cell, options)) _addCell(bounds, rIndex, cIndex);
        });
    });
}

function _getCellColIndex(sheet, ref) {
    const match = String(ref || '').match(/[A-Z]+/i);
    return match ? sheet.Utils.charToNum(match[0]) : null;
}

function _pendingCellHasContent(cell, options) {
    if (!cell) return false;
    const formulaNode = cell.f;
    if (options.formulas && formulaNode) return true;
    if (formulaNode) return false;
    if (options.contents) {
        if (cell.v !== undefined && cell.v !== null && cell.v !== '') return true;
        if (cell.is !== undefined && cell.is !== null && cell.is !== '') return true;
    }
    if (options.richText && cell.is?.r) return true;
    if (options.includeFormats && cell['_$s'] !== undefined) return true;
    return false;
}

function _addPendingCells(bounds, sheet, options) {
    const rowMap = sheet._pendingXlsxRowMap;
    if (!rowMap?.size) return;

    rowMap.forEach((rowXml, rIndex) => {
        const cells = Array.isArray(rowXml?.c) ? rowXml.c : (rowXml?.c ? [rowXml.c] : []);
        cells.forEach(cellXml => {
            if (!_pendingCellHasContent(cellXml, options)) return;
            const cIndex = _getCellColIndex(sheet, cellXml?.['_$r']);
            if (cIndex !== null) _addCell(bounds, rIndex, cIndex);
        });
    });
}

function _addRowColFormats(bounds, sheet) {
    sheet.rows.forEach((row, rIndex) => {
        if (!row) return;
        if (
            row._height !== sheet.defaultRowHeight ||
            row._hidden ||
            row._rowStyle ||
            row._outlineLevel > 0 ||
            row._collapsed
        ) {
            _addCell(bounds, rIndex, 0);
        }
    });
    sheet._pendingXlsxRowStateMap?.forEach((state, rIndex) => {
        if (
            state?.rowStyle ||
            state?.hidden ||
            state?.height !== sheet.defaultRowHeight ||
            state?.outlineLevel > 0 ||
            state?.collapsed
        ) {
            _addCell(bounds, rIndex, 0);
        }
    });
    sheet.cols.forEach((col, cIndex) => {
        if (col?._needBuild) _addCell(bounds, 0, cIndex);
    });
}

function _addComments(bounds, sheet) {
    if (!sheet.Comment?.getAll) return;
    sheet.Comment.getAll().forEach(comment => _addRef(bounds, sheet, comment.cellRef));
}

function _addDrawings(bounds, sheet) {
    if (!sheet.Drawing?.getAll) return;
    sheet.Drawing.getAll().forEach(drawing => _addRange(bounds, drawing.area));
}

function _addSparklines(bounds, sheet) {
    if (!sheet.Sparkline?.getAll) return;
    sheet.Sparkline.getAll().forEach(item => {
        _addCell(bounds, item.r, item.c);
        _addSqref(bounds, sheet, item.sqref);
    });
}

function _addConditionalFormats(bounds, sheet) {
    const rules = sheet.CF?.getAll?.() || sheet.CF?.rules || [];
    rules.forEach(rule => _addSqref(bounds, sheet, rule.sqref));
}

function _addPivotTables(bounds, sheet) {
    if (!sheet.PivotTable?.getAll) return;
    sheet.PivotTable.getAll().forEach(pivotTable => _addRef(bounds, sheet, pivotTable.location?.ref));
}

function _toRange(bounds, options, sheet) {
    if (bounds.maxR < 0 || bounds.maxC < 0 || !Number.isFinite(bounds.minR) || !Number.isFinite(bounds.minC)) {
        if (options.emptyAsNull) return null;
        const emptyRange = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
        return options.asString ? sheet.Utils.rangeNumToStr(emptyRange) : emptyRange;
    }

    const range = {
        s: { r: bounds.minR, c: bounds.minC },
        e: { r: bounds.maxR, c: bounds.maxC }
    };
    return options.asString ? sheet.Utils.rangeNumToStr(range) : range;
}

/**
 * Get the used range of the sheet without scanning the full grid.
 * @param {Object} [options]
 * @param {boolean} [options.contents=true] Include non-empty values.
 * @param {boolean} [options.formulas=true] Include formula cells.
 * @param {boolean} [options.richText=true] Include rich text cells.
 * @param {boolean} [options.hyperlinks=true] Include hyperlink cells.
 * @param {boolean} [options.dataValidations=true] Include data validation cells.
 * @param {boolean} [options.controls=true] Include cell control cells.
 * @param {boolean} [options.comments=true] Include comments.
 * @param {boolean} [options.merges=true] Include merged ranges.
 * @param {boolean} [options.tables=true] Include table ranges.
 * @param {boolean} [options.autoFilters=true] Include AutoFilter ranges.
 * @param {boolean} [options.drawings=false] Include drawing ranges.
 * @param {boolean} [options.sparklines=false] Include sparkline cells and source ranges.
 * @param {boolean} [options.conditionalFormats=false] Include conditional formatting ranges.
 * @param {boolean} [options.pivotTables=false] Include pivot table output ranges.
 * @param {boolean} [options.includeFormats=false] Include explicitly formatted cells.
 * @param {boolean} [options.includeRowColFormats=false] Include row and column format or visibility markers.
 * @param {boolean} [options.emptyAsNull=false] Return null when the sheet is empty.
 * @param {boolean} [options.asString=false] Return A1 range string instead of range object.
 * @returns {{s:{r:number,c:number},e:{r:number,c:number}}|string|null}
 */
export function getUsedRange(options = {}) {
    this._init?.();

    const usedOptions = _normalizeUsedRangeOptions(options);
    const bounds = _createBounds();

    _addExistingCells(bounds, this, usedOptions);
    _addPendingCells(bounds, this, usedOptions);

    if (usedOptions.includeRowColFormats) _addRowColFormats(bounds, this);
    if (usedOptions.merges) (this.merges || []).forEach(range => _addRange(bounds, range));
    if (usedOptions.tables) this.Table?.getAll?.().forEach(table => _addRange(bounds, table.range));
    if (usedOptions.autoFilters) this.AutoFilter?.getEnabledScopes?.().forEach(scope => _addRange(bounds, scope.range));
    if (usedOptions.comments) _addComments(bounds, this);
    if (usedOptions.drawings) _addDrawings(bounds, this);
    if (usedOptions.sparklines) _addSparklines(bounds, this);
    if (usedOptions.conditionalFormats) _addConditionalFormats(bounds, this);
    if (usedOptions.pivotTables) _addPivotTables(bounds, this);

    return _toRange(bounds, usedOptions, this);
}
