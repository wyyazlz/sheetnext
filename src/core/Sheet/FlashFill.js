const TOKEN_PATTERNS = [
    { key: 'nonSpace', regex: /[^\s]+/g },
    { key: 'alphaNum', regex: /[A-Za-z0-9]+/g },
    { key: 'alpha', regex: /[A-Za-z]+/g },
    { key: 'digit', regex: /\d+/g },
    { key: 'cjk', regex: /[\u4E00-\u9FFF]+/g }
];

const SPLIT_DELIMITERS = [' ', '-', '_', '/', '\\', '.', ',', ';', ':', '|', '@', '#'];
const REMOVE_CHAR_CANDIDATES = [' ', '-', '_', '/', '\\', '.', ',', ';', ':', '(', ')', '+'];

function _toText(value) {
    return value == null ? '' : String(value);
}

function _normalizeRange(sheet, range) {
    const fallback = sheet.activeAreas?.[0] ?? { s: sheet.activeCell, e: sheet.activeCell };
    const rawRange = range ?? fallback;
    const parsedRange = typeof rawRange === 'string' ? sheet.rangeStrToNum(rawRange) : rawRange;
    if (!parsedRange?.s || !parsedRange?.e) {
        throw new Error('Invalid flash fill target range.');
    }
    const normalized = {
        s: {
            r: Math.min(parsedRange.s.r, parsedRange.e.r),
            c: Math.min(parsedRange.s.c, parsedRange.e.c)
        },
        e: {
            r: Math.max(parsedRange.s.r, parsedRange.e.r),
            c: Math.max(parsedRange.s.c, parsedRange.e.c)
        }
    };
    if (normalized.s.c !== normalized.e.c) {
        throw new Error('Flash Fill only supports a single-column target range.');
    }
    return normalized;
}

function _normalizeSourceColumn(sheet, targetRange, options = {}) {
    if (options.sourceCol !== undefined && options.sourceCol !== null) {
        if (typeof options.sourceCol === 'string') {
            const col = sheet.Utils.charToNum(options.sourceCol.replace(/\$/g, '').trim().toUpperCase());
            if (Number.isInteger(col) && col >= 0) return col;
            throw new Error('Invalid source column.');
        }
        if (Number.isInteger(options.sourceCol) && options.sourceCol >= 0) {
            return options.sourceCol;
        }
        throw new Error('Invalid source column.');
    }

    const targetCol = targetRange.s.c;
    const candidates = [];
    if (targetCol > 0) candidates.push(targetCol - 1);
    if (targetCol + 1 < sheet.colCount) candidates.push(targetCol + 1);
    if (candidates.length === 0) {
        throw new Error('No adjacent source column was found for Flash Fill.');
    }
    if (candidates.length === 1) return candidates[0];

    const countNonEmpty = (col) => {
        let count = 0;
        for (let r = targetRange.s.r; r <= targetRange.e.r; r++) {
            const text = _toText(sheet.getCell(r, col).showVal);
            if (text !== '') count++;
        }
        return count;
    };

    const left = candidates[0];
    const right = candidates[1];
    const leftCount = countNonEmpty(left);
    const rightCount = countNonEmpty(right);
    return rightCount > leftCount ? right : left;
}

function _splitByDelimiter(text, delimiter) {
    if (delimiter === 'whitespace') {
        return _toText(text).trim().split(/\s+/).filter(Boolean);
    }
    return _toText(text).split(delimiter).filter(part => part !== '');
}

function _pickByIndex(list, indexMode, indexValue) {
    if (!Array.isArray(list) || list.length === 0) return '';
    if (indexMode === 'start') {
        return list[indexValue] ?? '';
    }
    const pos = list.length - indexValue;
    return pos >= 0 && pos < list.length ? list[pos] : '';
}

function _titleCase(text) {
    return _toText(text).replace(/[A-Za-z]+/g, part => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase());
}

function _buildCandidates(sampleSource, sampleTarget, allExamples) {
    const source = _toText(sampleSource);
    const target = _toText(sampleTarget);
    const candidates = [];
    const seen = new Set();

    const pushCandidate = (id, score, apply) => {
        if (seen.has(id)) return;
        seen.add(id);
        candidates.push({ id, score, apply });
    };

    // Exact source transforms
    if (target === source) pushCandidate('transform:identity', 15, (text) => _toText(text));
    if (target === source.toUpperCase()) pushCandidate('transform:upper', 20, (text) => _toText(text).toUpperCase());
    if (target === source.toLowerCase()) pushCandidate('transform:lower', 20, (text) => _toText(text).toLowerCase());
    if (target === _titleCase(source)) pushCandidate('transform:title', 18, (text) => _titleCase(text));
    if (target === source.replace(/\D+/g, '')) pushCandidate('transform:digits-only', 22, (text) => _toText(text).replace(/\D+/g, ''));

    // Remove fixed characters
    REMOVE_CHAR_CANDIDATES.forEach(ch => {
        if (target !== source.split(ch).join('')) return;
        pushCandidate(`remove-char:${JSON.stringify(ch)}`, 24, (text) => _toText(text).split(ch).join(''));
    });

    // Regex token extraction
    TOKEN_PATTERNS.forEach(pattern => {
        const parts = source.match(pattern.regex) || [];
        parts.forEach((part, idx) => {
            if (part !== target) return;
            pushCandidate(`token:${pattern.key}:start:${idx}`, 40, (text) => {
                const segments = _toText(text).match(pattern.regex) || [];
                return _pickByIndex(segments, 'start', idx);
            });
            const fromEnd = parts.length - idx;
            pushCandidate(`token:${pattern.key}:end:${fromEnd}`, 45, (text) => {
                const segments = _toText(text).match(pattern.regex) || [];
                return _pickByIndex(segments, 'end', fromEnd);
            });
        });
    });

    // Delimiter split extraction
    const delimiterDefs = ['whitespace', ...SPLIT_DELIMITERS];
    delimiterDefs.forEach(delimiter => {
        const parts = _splitByDelimiter(source, delimiter);
        parts.forEach((part, idx) => {
            if (part !== target) return;
            pushCandidate(`split:${JSON.stringify(delimiter)}:start:${idx}`, 50, (text) => {
                const segments = _splitByDelimiter(text, delimiter);
                return _pickByIndex(segments, 'start', idx);
            });
            const fromEnd = parts.length - idx;
            pushCandidate(`split:${JSON.stringify(delimiter)}:end:${fromEnd}`, 55, (text) => {
                const segments = _splitByDelimiter(text, delimiter);
                return _pickByIndex(segments, 'end', fromEnd);
            });
        });
    });

    // Fixed substring
    if (target !== '') {
        let from = 0;
        while (from < source.length) {
            const start = source.indexOf(target, from);
            if (start === -1) break;
            const len = target.length;
            pushCandidate(`substr:start:${start}:len:${len}`, 30, (text) => {
                const t = _toText(text);
                if (start + len > t.length) return '';
                return t.slice(start, start + len);
            });
            const fromEnd = source.length - start - len;
            pushCandidate(`substr:end:${fromEnd}:len:${len}`, 32, (text) => {
                const t = _toText(text);
                const realStart = t.length - fromEnd - len;
                if (realStart < 0 || realStart + len > t.length) return '';
                return t.slice(realStart, realStart + len);
            });
            from = start + 1;
        }
    }

    // Lookup fallback for repeated keys
    const lookupMap = new Map();
    let lookupConsistent = true;
    allExamples.forEach(example => {
        if (!lookupMap.has(example.source)) {
            lookupMap.set(example.source, example.target);
            return;
        }
        if (lookupMap.get(example.source) !== example.target) {
            lookupConsistent = false;
        }
    });
    if (lookupConsistent && lookupMap.size > 0) {
        pushCandidate('lookup-map', 10, (text) => lookupMap.get(_toText(text)) ?? '');
    }

    // Prefix/suffix wrapper based on extractors
    const baseCandidates = candidates.slice();
    baseCandidates.forEach(base => {
        const extracted = _toText(base.apply(source));
        if (!extracted || extracted === target) return;
        const idx = target.indexOf(extracted);
        if (idx === -1) return;
        const prefix = target.slice(0, idx);
        const suffix = target.slice(idx + extracted.length);
        pushCandidate(`wrap:${base.id}:${JSON.stringify(prefix)}:${JSON.stringify(suffix)}`, base.score + 2, (text) => {
            const core = _toText(base.apply(text));
            if (core === '') return '';
            return prefix + core + suffix;
        });
    });

    return candidates;
}

function _resolveStrategy(examples) {
    if (!examples.length) return null;
    const first = examples[0];
    const candidates = _buildCandidates(first.source, first.target, examples);
    let best = null;
    for (const candidate of candidates) {
        const valid = examples.every(example => _toText(candidate.apply(example.source)) === example.target);
        if (!valid) continue;
        if (!best || candidate.score > best.score) {
            best = candidate;
        }
    }
    return best;
}

export function flashFill(range, options = {}) {
    const targetRange = _normalizeRange(this, range);
    const sourceCol = _normalizeSourceColumn(this, targetRange, options);
    const targetCol = targetRange.s.c;
    const overwrite = options.overwrite === true;

    const rows = [];
    for (let r = targetRange.s.r; r <= targetRange.e.r; r++) {
        const sourceText = _toText(this.getCell(r, sourceCol).showVal);
        const targetText = _toText(this.getCell(r, targetCol).editVal);
        rows.push({ row: r, source: sourceText, target: targetText });
    }

    const examples = rows.filter(item => item.source !== '' && item.target !== '');
    if (examples.length === 0) {
        throw new Error('Please provide at least one example value in the target column before running Flash Fill.');
    }

    const strategy = _resolveStrategy(examples);
    if (!strategy) {
        throw new Error('Unable to infer a Flash Fill pattern from the provided examples.');
    }

    const sourceRange = {
        s: { r: targetRange.s.r, c: sourceCol },
        e: { r: targetRange.e.r, c: sourceCol }
    };
    const predictions = rows.map(item => ({
        ...item,
        predicted: _toText(strategy.apply(item.source))
    }));

    const summary = {
        success: true,
        targetRange,
        sourceRange,
        sourceCol,
        exampleCount: examples.length,
        strategy: strategy.id,
        filledCount: 0
    };
    if (options.previewOnly) {
        return { ...summary, predictions };
    }

    const beforeEvent = this.SN.Event.emit('beforeFlashFill', {
        sheet: this,
        targetRange,
        sourceRange,
        strategy: strategy.id,
        exampleCount: examples.length
    });
    if (beforeEvent.canceled) {
        return { ...summary, success: false, canceled: true };
    }

    return this.SN._withOperation('flashFill', {
        sheet: this,
        targetRange,
        sourceRange,
        strategy: strategy.id
    }, () => {
        let filledCount = 0;

        predictions.forEach(item => {
            if (!overwrite && item.target !== '') return;
            if (item.predicted === '') return;
            if (item.predicted === item.target) return;
            this.getCell(item.row, targetCol).editVal = item.predicted;
            filledCount++;
        });

        this.SN._recordChange({
            type: 'flashFill',
            sheet: this,
            targetRange,
            sourceRange,
            strategy: strategy.id,
            exampleCount: examples.length,
            filledCount
        });

        this.SN.Event.emit('afterFlashFill', {
            sheet: this,
            targetRange,
            sourceRange,
            strategy: strategy.id,
            exampleCount: examples.length,
            filledCount
        });

        return { ...summary, filledCount };
    });
}
