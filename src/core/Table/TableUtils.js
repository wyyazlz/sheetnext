/**
 * Table Tool Functions
 */

/**
 * Table Tool Functions
 */

/**
 * Parsing range string "A1:H14" - {s: {r, c}, e: {r, c}
 * @param {string} refStr
 * @returns {Object}
 */
export function parseRef(refStr) {
    if (!refStr || typeof refStr !== 'string') return null;

    const [start, end] = refStr.split(':');
    if (!start) return null;

    const s = parseCellRef(start);
    const e = end ? parseCellRef(end) : { ...s };

    return { s, e };
}

/**
 * Range Object String
 * @param {Object} ref - { s: {r, c}, e: {r, c} }
 * @returns {string}
 */
export function toRef(ref) {
    if (!ref || !ref.s) return '';
    const start = toCellRef(ref.s.r, ref.s.c);
    const end = toCellRef(ref.e.r, ref.e.c);
    return start === end ? start : `${start}:${end}`;
}

/**
 * Parsing cell reference "A1" - {r, c}
 * @param {string} cellRef
 * @returns {Object}
 */
export function parseCellRef(cellRef) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return { r: 0, c: 0 };

    const col = colToIndex(match[1]);
    const row = parseInt(match[2], 10) - 1;

    return { r: row, c: col };
}

/**
 * Row Indexing Cell Reference
 * @param {number} row - 0-based
 * @param {number} col - 0-based
 * @returns {string}
 */
export function toCellRef(row, col) {
    return indexToCol(col) + (row + 1);
}

/**
 * "A" - 0, "Z" - 25, "AA" - 26
 * @param {string} col
 * @returns {number}
 */
export function colToIndex(col) {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
        index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index - 1;
}

/**
 * Index letters 0- > "A", 25- > "Z", 26- > "AA"
 * @param {number} index - 0-based
 * @returns {string}
 */
export function indexToCol(index) {
    let col = '';
    let n = index + 1;
    while (n > 0) {
        n--;
        col = String.fromCharCode(65 + (n % 26)) + col;
        n = Math.floor(n / 26);
    }
    return col;
}

/**
 * Generate unique table name
 * @param {Map} tables - Existing Table
 * @param {string} baseName - Base Name
 * @returns {string}
 */
export function generateTableName(tables, baseName = '表') {
    let index = 1;
    let name = `${baseName}${index}`;

    const existingNames = new Set();
    tables.forEach(table => existingNames.add(table.name));

    while (existingNames.has(name)) {
        index++;
        name = `${baseName}${index}`;
    }

    return name;
}
