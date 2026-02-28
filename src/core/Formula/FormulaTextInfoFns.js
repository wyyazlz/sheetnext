export function createTextInfoFns(ctx) {
    const {
        stack,
        gVal,
        execute,
        getFlatValues,
        splitByDelims,
        toDelims,
        toArrayText,
        escapeRegex,
        ensureInteger,
        halfWidthText,
        fullWidthText,
        firstCellFromArg,
        currentCell,
        SN
    } = ctx;

    return {
        REPLACEB: () => execute('REPLACE', stack),
        FINDB: () => execute('FIND', stack),
        SEARCHB: () => execute('SEARCH', stack),
        TEXTJOIN: () => {
            const delimiter = String(gVal(stack[0]) ?? '');
            const ignoreEmpty = !!gVal(stack[1]);
            const values = [];
            for (let i = 2; i < stack.length; i++) {
                getFlatValues(stack[i]).forEach(v => {
                    if (ignoreEmpty && (v === '' || v === null || v === undefined)) return;
                    values.push(String(v ?? ''));
                });
            }
            return values.join(delimiter);
        },
        TEXTBEFORE: () => {
            const text = String(gVal(stack[0]) ?? '');
            const delimiter = String(gVal(stack[1]) ?? '');
            const instanceRaw = stack[2] !== undefined ? Number(gVal(stack[2])) : 1;
            const matchMode = stack[3] !== undefined ? Number(gVal(stack[3])) : 0;
            const matchEnd = stack[4] !== undefined ? !!gVal(stack[4]) : false;
            const fallback = stack[5] !== undefined ? gVal(stack[5]) : null;
            const instance = instanceRaw === 0 ? 1 : Math.trunc(instanceRaw);

            if (delimiter === '') return instance > 0 ? '' : text;

            const source = matchMode === 1 ? text.toLowerCase() : text;
            const needle = matchMode === 1 ? delimiter.toLowerCase() : delimiter;
            let idx = -1;

            if (instance > 0) {
                let from = 0;
                for (let i = 0; i < instance; i++) {
                    idx = source.indexOf(needle, from);
                    if (idx === -1) break;
                    from = idx + needle.length;
                }
            } else {
                let from = source.length;
                for (let i = 0; i < Math.abs(instance); i++) {
                    idx = source.lastIndexOf(needle, from - 1);
                    if (idx === -1) break;
                    from = idx;
                }
            }

            if (idx === -1) {
                if (fallback !== null) return fallback;
                if (matchEnd) return text;
                throw new Error('#N/A');
            }
            return text.substring(0, idx);
        },
        TEXTAFTER: () => {
            const text = String(gVal(stack[0]) ?? '');
            const delimiter = String(gVal(stack[1]) ?? '');
            const instanceRaw = stack[2] !== undefined ? Number(gVal(stack[2])) : 1;
            const matchMode = stack[3] !== undefined ? Number(gVal(stack[3])) : 0;
            const matchEnd = stack[4] !== undefined ? !!gVal(stack[4]) : false;
            const fallback = stack[5] !== undefined ? gVal(stack[5]) : null;
            const instance = instanceRaw === 0 ? 1 : Math.trunc(instanceRaw);

            if (delimiter === '') return instance > 0 ? text : '';

            const source = matchMode === 1 ? text.toLowerCase() : text;
            const needle = matchMode === 1 ? delimiter.toLowerCase() : delimiter;
            let idx = -1;

            if (instance > 0) {
                let from = 0;
                for (let i = 0; i < instance; i++) {
                    idx = source.indexOf(needle, from);
                    if (idx === -1) break;
                    from = idx + needle.length;
                }
            } else {
                let from = source.length;
                for (let i = 0; i < Math.abs(instance); i++) {
                    idx = source.lastIndexOf(needle, from - 1);
                    if (idx === -1) break;
                    from = idx;
                }
            }

            if (idx === -1) {
                if (fallback !== null) return fallback;
                if (matchEnd) return '';
                throw new Error('#N/A');
            }
            return text.substring(idx + delimiter.length);
        },
        TEXTSPLIT: () => {
            const text = String(gVal(stack[0]) ?? '');
            const colDelRaw = gVal(stack[1]);
            const rowDelRaw = stack[2] !== undefined ? gVal(stack[2]) : null;
            const ignoreEmpty = stack[3] !== undefined ? !!gVal(stack[3]) : false;
            const matchMode = stack[4] !== undefined ? Number(gVal(stack[4])) : 0;
            const padWith = stack[5] !== undefined ? gVal(stack[5]) : '';

            const rowTokens = rowDelRaw == null ? [text] : splitByDelims(text, toDelims(rowDelRaw), matchMode);
            const rows = rowTokens.map(rowText => splitByDelims(rowText, toDelims(colDelRaw), matchMode));
            const normalized = rows.map(row => ignoreEmpty ? row.filter(x => x !== '') : row);
            const maxCols = normalized.reduce((m, row) => Math.max(m, row.length), 0);
            return normalized.map(row => row.length < maxCols ? [...row, ...Array(maxCols - row.length).fill(padWith)] : row);
        },
        ARRAYTOTEXT: () => toArrayText(gVal(stack[0])),
        NUMBERVALUE: () => {
            const text = String(gVal(stack[0]) ?? '').trim();
            const decimalSep = stack[1] !== undefined ? String(gVal(stack[1])) : '.';
            const groupSep = stack[2] !== undefined ? String(gVal(stack[2])) : ',';
            let normalized = text;
            if (groupSep) normalized = normalized.replace(new RegExp(escapeRegex(groupSep), 'g'), '');
            if (decimalSep && decimalSep !== '.') normalized = normalized.replace(new RegExp(escapeRegex(decimalSep), 'g'), '.');
            const out = Number(normalized);
            if (!Number.isFinite(out)) throw new Error('#VALUE!');
            return out;
        },
        ENCODEURL: () => encodeURIComponent(String(gVal(stack[0]) ?? '')),
        UNICHAR: () => {
            const code = ensureInteger(gVal(stack[0]));
            if (code < 1 || code > 1114111) throw new Error('#VALUE!');
            return String.fromCodePoint(code);
        },
        UNICODE: () => {
            const text = String(gVal(stack[0]) ?? '');
            if (!text) throw new Error('#VALUE!');
            return text.codePointAt(0);
        },
        VALUETOTEXT: () => String(gVal(stack[0]) ?? ''),
        REGEXTEST: () => {
            const text = String(gVal(stack[0]) ?? '');
            const pattern = String(gVal(stack[1]) ?? '');
            const flags = stack[2] ? String(gVal(stack[2])) : '';
            return new RegExp(pattern, flags).test(text);
        },
        REGEXEXTRACT: () => {
            const text = String(gVal(stack[0]) ?? '');
            const pattern = String(gVal(stack[1]) ?? '');
            const flags = stack[2] ? String(gVal(stack[2])) : '';
            const m = text.match(new RegExp(pattern, flags));
            if (!m) throw new Error('#N/A');
            return m[1] ?? m[0];
        },
        REGEXREPLACE: () => {
            const text = String(gVal(stack[0]) ?? '');
            const pattern = String(gVal(stack[1]) ?? '');
            const replacement = String(gVal(stack[2]) ?? '');
            const flags = stack[3] ? String(gVal(stack[3])) : 'g';
            return text.replace(new RegExp(pattern, flags), replacement);
        },
        ASC: () => halfWidthText(gVal(stack[0])),
        JIS: () => fullWidthText(gVal(stack[0])),
        'ERROR.TYPE': () => execute('ERROR_TYPE', stack),
        ISFORMULA: () => {
            const cell = firstCellFromArg(stack[0]);
            return !!(cell && typeof cell.editVal === 'string' && cell.editVal.startsWith('='));
        },
        SHEET: () => {
            if (!stack.length) {
                const sheet = currentCell?.row?.sheet || SN.activeSheet;
                const index = SN.sheets.findIndex(s => s.name === sheet?.name);
                return index >= 0 ? index + 1 : 1;
            }
            const cell = firstCellFromArg(stack[0]);
            if (cell?.row?.sheet) {
                const index = SN.sheets.findIndex(s => s.name === cell.row.sheet.name);
                if (index >= 0) return index + 1;
            }
            const raw = String(gVal(stack[0]) ?? '');
            const match = raw.match(/^'?([^']*)'?!/);
            const name = match ? match[1] : raw;
            const index = SN.sheets.findIndex(s => s.name === name);
            if (index < 0) throw new Error('#N/A');
            return index + 1;
        },
        SHEETS: () => {
            if (!stack.length) return SN.sheets.length;
            return 1;
        },
        CELL: () => {
            const infoType = String(gVal(stack[0]) ?? '').toLowerCase();
            const refCell = stack[1] !== undefined ? (firstCellFromArg(stack[1]) || currentCell) : currentCell;
            if (!refCell) throw new Error('#VALUE!');
            switch (infoType) {
                case 'address':
                    return `$${SN.Utils.numToChar(refCell.cIndex)}$${refCell.row.rIndex + 1}`;
                case 'row':
                    return refCell.row.rIndex + 1;
                case 'col':
                    return refCell.cIndex + 1;
                case 'contents':
                    return refCell.calcVal;
                case 'type':
                    if (refCell.calcVal == null || refCell.calcVal === '') return 'b';
                    if (typeof refCell.calcVal === 'string') return 'l';
                    return 'v';
                case 'filename':
                    return '';
                default:
                    throw new Error('#VALUE!');
            }
        },
        INFO: () => {
            const infoType = String(gVal(stack[0]) ?? '').toLowerCase();
            switch (infoType) {
                case 'recalc': return 'Automatic';
                case 'origin': return '$A$1';
                case 'numfile': return 1;
                case 'release': return 'SheetNext';
                case 'system': return 'pcdos';
                case 'osversion': return 'SheetNext';
                default: throw new Error('#N/A');
            }
        },
    };
}

export default createTextInfoFns;
