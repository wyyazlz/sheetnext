const DELIMITER_ALIAS_MAP = {
    tab: '\t',
    '\\t': '\t',
    comma: ',',
    semicolon: ';',
    space: ' '
};

function _normalizeRange(sheet, range) {
    const fallback = sheet.activeAreas?.[0] ?? { s: sheet.activeCell, e: sheet.activeCell };
    const rawRange = range ?? fallback;
    const parsedRange = typeof rawRange === 'string' ? sheet.rangeStrToNum(rawRange) : rawRange;
    if (!parsedRange?.s || !parsedRange?.e) {
        throw new Error('Invalid source range.');
    }
    return {
        s: {
            r: Math.min(parsedRange.s.r, parsedRange.e.r),
            c: Math.min(parsedRange.s.c, parsedRange.e.c)
        },
        e: {
            r: Math.max(parsedRange.s.r, parsedRange.e.r),
            c: Math.max(parsedRange.s.c, parsedRange.e.c)
        }
    };
}

function _normalizeDestination(sheet, sourceRange, destination) {
    if (!destination) return { r: sourceRange.s.r, c: sourceRange.s.c };
    const parsed = typeof destination === 'string'
        ? sheet.Utils.cellStrToNum(destination.replace(/\$/g, '').trim())
        : destination;
    if (!parsed || parsed.r < 0 || parsed.c < 0) {
        throw new Error('Invalid destination.');
    }
    return { r: parsed.r, c: parsed.c };
}

function _normalizeDelimiters(options = {}) {
    let rawDelimiters = options.delimiters ?? [];
    if (!Array.isArray(rawDelimiters)) {
        rawDelimiters = [rawDelimiters];
    }
    if (options.delimiter !== undefined && options.delimiter !== null) {
        rawDelimiters.push(options.delimiter);
    }

    const delimiterSet = new Set();
    rawDelimiters.forEach(item => {
        if (item === undefined || item === null) return;
        const normalized = DELIMITER_ALIAS_MAP[item] ?? String(item);
        if (!normalized) return;
        for (const ch of normalized) delimiterSet.add(ch);
    });
    return Array.from(delimiterSet);
}

function _normalizeTextQualifier(textQualifier) {
    if (textQualifier === null || textQualifier === undefined) return '"';
    if (textQualifier === '' || textQualifier === 'none') return null;
    return String(textQualifier)[0] ?? null;
}

function _splitTextByDelimiters(text, delimiters, options = {}) {
    const source = text == null ? '' : String(text);
    const textQualifier = options.textQualifier ?? '"';
    const mergeConsecutive = options.treatConsecutiveDelimitersAsOne === true;
    const delimiterSet = new Set(delimiters);
    if (!delimiterSet.size) return [source];

    const tokens = [];
    let token = '';
    let inQualifiedText = false;

    for (let i = 0; i < source.length; i++) {
        const ch = source[i];
        if (textQualifier && ch === textQualifier) {
            if (inQualifiedText && source[i + 1] === textQualifier) {
                token += textQualifier;
                i++;
            } else {
                inQualifiedText = !inQualifiedText;
            }
            continue;
        }

        const isDelimiter = !inQualifiedText && delimiterSet.has(ch);
        if (isDelimiter) {
            tokens.push(token);
            token = '';
            if (mergeConsecutive) {
                while (i + 1 < source.length && delimiterSet.has(source[i + 1])) i++;
            }
            continue;
        }
        token += ch;
    }

    tokens.push(token);
    return tokens;
}

export function textToColumns(range, options = {}) {
    const sourceRange = _normalizeRange(this, range);
    if (sourceRange.s.c !== sourceRange.e.c) {
        throw new Error('Text to Columns only supports a single-column source range.');
    }

    const delimiters = _normalizeDelimiters(options);
    if (delimiters.length === 0) {
        throw new Error('At least one delimiter is required.');
    }

    const destination = _normalizeDestination(this, sourceRange, options.destination);
    const textQualifier = _normalizeTextQualifier(options.textQualifier);
    const splitOptions = {
        textQualifier,
        treatConsecutiveDelimitersAsOne: options.treatConsecutiveDelimitersAsOne === true
    };

    const rows = [];
    for (let r = sourceRange.s.r; r <= sourceRange.e.r; r++) {
        const sourceCell = this.getCell(r, sourceRange.s.c);
        rows.push(String(sourceCell.showVal ?? ''));
    }

    const splitRows = rows.map(text => _splitTextByDelimiters(text, delimiters, splitOptions));
    const columnCount = splitRows.reduce((max, parts) => Math.max(max, parts.length), 1);
    const rowCount = splitRows.length;
    const targetRange = {
        s: { r: destination.r, c: destination.c },
        e: { r: destination.r + rowCount - 1, c: destination.c + columnCount - 1 }
    };

    const summary = {
        success: true,
        sourceRange,
        targetRange,
        rowCount,
        columnCount
    };

    if (options.previewOnly) return summary;

    const beforeEvent = this.SN.Event.emit('beforeTextToColumns', {
        sheet: this,
        sourceRange,
        targetRange,
        options: {
            delimiters,
            textQualifier,
            treatConsecutiveDelimitersAsOne: splitOptions.treatConsecutiveDelimitersAsOne
        }
    });
    if (beforeEvent.canceled) {
        return { ...summary, success: false, canceled: true };
    }

    return this.SN._withOperation('textToColumns', {
        sheet: this,
        sourceRange,
        targetRange
    }, () => {
        splitRows.forEach((parts, rowOffset) => {
            const targetRow = destination.r + rowOffset;
            for (let colOffset = 0; colOffset < columnCount; colOffset++) {
                const targetCol = destination.c + colOffset;
                this.getCol(targetCol);
                const targetCell = this.getCell(targetRow, targetCol);
                targetCell.editVal = parts[colOffset] ?? '';
            }
        });

        this.SN._recordChange({
            type: 'textToColumns',
            sheet: this,
            sourceRange,
            targetRange,
            options: {
                delimiters,
                textQualifier,
                treatConsecutiveDelimitersAsOne: splitOptions.treatConsecutiveDelimitersAsOne
            },
            rowCount,
            columnCount
        });

        this.SN.Event.emit('afterTextToColumns', {
            sheet: this,
            sourceRange,
            targetRange,
            rowCount,
            columnCount
        });

        return summary;
    });
}
