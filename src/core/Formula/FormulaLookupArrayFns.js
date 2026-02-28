export function createLookupArrayFns(ctx) {
    const {
        stack,
        gVal,
        ensureInteger,
        toMatrixValue,
        matchIndex,
        firstCellFromArg
    } = ctx;

    return {
        ROWS: () => {
            const value = gVal(stack[0]);
            if (Array.isArray(value)) return Array.isArray(value[0]) ? value.length : 1;
            return 1;
        },
        COLUMNS: () => {
            const value = gVal(stack[0]);
            if (Array.isArray(value)) return Array.isArray(value[0]) ? value[0].length : value.length;
            return 1;
        },
        CHOOSE: () => {
            const idx = ensureInteger(gVal(stack[0]), '#VALUE!');
            if (idx < 1 || idx >= stack.length) throw new Error('#VALUE!');
            return gVal(stack[idx]);
        },
        HYPERLINK: () => {
            const link = String(gVal(stack[0]) ?? '');
            const friendly = stack[1] !== undefined ? gVal(stack[1]) : null;
            return friendly ?? link;
        },
        XMATCH: () => {
            const lookupValue = gVal(stack[0]);
            const lookupArray = gVal(stack[1]);
            const values = Array.isArray(lookupArray)
                ? (Array.isArray(lookupArray[0]) ? lookupArray.flat() : lookupArray)
                : [lookupArray];
            const matchMode = stack[2] !== undefined ? Number(gVal(stack[2])) : 0;
            const searchMode = stack[3] !== undefined ? Number(gVal(stack[3])) : 1;
            const idx = matchIndex(values, lookupValue, matchMode, searchMode);
            if (idx < 0) throw new Error('#N/A');
            return idx + 1;
        },
        FORMULATEXT: () => {
            const cell = firstCellFromArg(stack[0]);
            if (!cell || typeof cell.editVal !== 'string' || !cell.editVal.startsWith('=')) throw new Error('#N/A');
            return cell.editVal;
        },
        TAKE: () => {
            const matrix = toMatrixValue(stack[0]);
            const rowRaw = stack[1] !== undefined ? ensureInteger(gVal(stack[1])) : matrix.length;
            const colRaw = stack[2] !== undefined ? ensureInteger(gVal(stack[2])) : (matrix[0]?.length || 0);
            if (rowRaw === 0 || colRaw === 0) throw new Error('#VALUE!');

            const rowCount = Math.min(Math.abs(rowRaw), matrix.length);
            const rows = rowRaw > 0 ? matrix.slice(0, rowCount) : matrix.slice(matrix.length - rowCount);
            const colCount = rows.length ? Math.min(Math.abs(colRaw), rows[0].length) : 0;
            return rows.map(row => colRaw > 0 ? row.slice(0, colCount) : row.slice(row.length - colCount));
        },
        DROP: () => {
            const matrix = toMatrixValue(stack[0]);
            const rowRaw = stack[1] !== undefined ? ensureInteger(gVal(stack[1])) : 0;
            const colRaw = stack[2] !== undefined ? ensureInteger(gVal(stack[2])) : 0;

            const rowKeep = Math.max(matrix.length - Math.min(Math.abs(rowRaw), matrix.length), 0);
            const rowStart = rowRaw >= 0 ? Math.min(rowRaw, matrix.length) : 0;
            const rowEnd = rowRaw >= 0 ? matrix.length : rowKeep;
            const rows = matrix.slice(rowStart, rowEnd);
            if (!rows.length) return [[]];

            const totalCols = rows[0].length;
            const colKeep = Math.max(totalCols - Math.min(Math.abs(colRaw), totalCols), 0);
            const colStart = colRaw >= 0 ? Math.min(colRaw, totalCols) : 0;
            const colEnd = colRaw >= 0 ? totalCols : colKeep;
            const out = rows.map(row => row.slice(colStart, colEnd));
            return out.length ? out : [[]];
        },
        EXPAND: () => {
            const matrix = toMatrixValue(stack[0]);
            const targetRows = ensureInteger(gVal(stack[1]));
            const targetCols = stack[2] !== undefined ? ensureInteger(gVal(stack[2])) : (matrix[0]?.length || 0);
            const padWith = stack[3] !== undefined ? gVal(stack[3]) : '#N/A';
            if (targetRows < matrix.length || targetCols < (matrix[0]?.length || 0)) throw new Error('#VALUE!');

            return Array.from({ length: targetRows }, (_, r) =>
                Array.from({ length: targetCols }, (_, c) =>
                    r < matrix.length && c < matrix[0].length ? matrix[r][c] : padWith
                )
            );
        },
        TRIMRANGE: () => {
            const matrix = toMatrixValue(stack[0]);
            const isBlank = (v) => v === null || v === undefined || v === '';
            let top = 0, bottom = matrix.length - 1;
            let left = 0, right = matrix[0].length - 1;

            while (top <= bottom && matrix[top].every(isBlank)) top++;
            while (bottom >= top && matrix[bottom].every(isBlank)) bottom--;

            const colBlank = (col) => {
                for (let r = top; r <= bottom; r++) {
                    if (!isBlank(matrix[r][col])) return false;
                }
                return true;
            };

            while (left <= right && colBlank(left)) left++;
            while (right >= left && colBlank(right)) right--;

            if (top > bottom || left > right) return [[""]];
            return matrix.slice(top, bottom + 1).map(row => row.slice(left, right + 1));
        },
    };
}

export default createLookupArrayFns;
