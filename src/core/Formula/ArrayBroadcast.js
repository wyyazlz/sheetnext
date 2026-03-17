/**
 * Array Broadcast Engine
 * Implement Excel-style array operations and broadcast mechanisms
 *
 * Broadcast Rules:
 * - Scalar expands to the same → dimension as any array scalar
 * - 1 × n row vectors and m × 1 column vectors are → extended to m × n matrices
 * - 1 × n row vectors and m × n matrix → row vectors copied m times
 * - m × 1 column vector copied n times with m × n matrix → column vector
 */

/**
 * Array Broadcast Engine
 * Implement Excel-style array operations and broadcast mechanisms
 *
 * Broadcast Rules:
 * - Scalar expands to the same → dimension as any array scalar
 * - 1 × n row vectors and m × 1 column vectors are → extended to m × n matrices
 * - 1 × n row vectors and m × n matrix → row vectors copied m times
 * - m × 1 column vector copied n times with m × n matrix → column vector
 */

/**
 * Gets the shape of the array [rows, cols]
 * @param {*} arr - Input values (scalar, one-dimensional array, two-dimensional array)
 * @returns {[number, number]} - [number of rows, number of columns]
 */
export function shape(arr) {
    if (!Array.isArray(arr)) return [1, 1];  // 标量
    if (arr.length === 0) return [0, 0];
    if (!Array.isArray(arr[0])) return [1, arr.length];  // 一维数组视为行向量
    return [arr.length, arr[0].length];  // 二维数组
}

/**
 * Determine if it is a two-dimensional array
 */
export function is2DArray(arr) {
    return Array.isArray(arr) && arr.length > 0 && Array.isArray(arr[0]);
}

/**
 * Normalize any value to a two-dimensional array
 * @param {*} val - Input value
 * @returns {Array<Array>} - 2-D Array
 */
export function to2D(val) {
    if (!Array.isArray(val)) return [[val]];  // 标量 → 1×1
    if (!Array.isArray(val[0])) return [val];  // 一维 → 1×n 行向量
    return val;  // 已经是二维
}

/**
 * Expand the 2D array to the specified dimension
 * @param {Array<Array>} arr - 2-D Array
 * @param {number} rows - Target Rows
 * @param {number} cols - Target number of columns
 * @returns {Array<Array>} - Expanded array
 */
export function expand(arr, rows, cols) {
    const [srcRows, srcCols] = shape(arr);
    const src = to2D(arr);

    // 已经是目标维度
    if (srcRows === rows && srcCols === cols) return src;

    // 验证广播兼容性
    if (srcRows !== 1 && srcRows !== rows) {
        throw new Error(`#VALUE! 无法广播: ${srcRows}行 → ${rows}行`);
    }
    if (srcCols !== 1 && srcCols !== cols) {
        throw new Error(`#VALUE! 无法广播: ${srcCols}列 → ${cols}列`);
    }

    const result = [];
    for (let i = 0; i < rows; i++) {
        const row = [];
        const srcRowIdx = srcRows === 1 ? 0 : i;
        for (let j = 0; j < cols; j++) {
            const srcColIdx = srcCols === 1 ? 0 : j;
            row.push(src[srcRowIdx][srcColIdx]);
        }
        result.push(row);
    }
    return result;
}

/**
 * Broadcast two arrays to the same dimension
 * @param {*} a - First operand
 * @param {*} b - Second operand
 * @returns {[Array<Array>, Array<Array>]} - Expanded two arrays
 */
export function broadcast(a, b) {
    const [rowsA, colsA] = shape(a);
    const [rowsB, colsB] = shape(b);

    // 计算目标维度
    const rows = Math.max(rowsA, rowsB);
    const cols = Math.max(colsA, colsB);

    return [
        expand(a, rows, cols),
        expand(b, rows, cols)
    ];
}

/**
 * Element-by-element operation on two arrays (broadcast support)
 * @param {*} a - First operand
 * @param {*} b - Second operand
 * @param {Function} op - Operating function (a, b) = > result
 * @returns {*} - Result of the operation (scalar or two-dimensional array)
 */
export function elementWise(a, b, op) {
    const aIs2D = is2DArray(a) || (Array.isArray(a) && a.length > 1);
    const bIs2D = is2DArray(b) || (Array.isArray(b) && b.length > 1);

    // 两个都是标量，返回标量
    if (!Array.isArray(a) && !Array.isArray(b)) {
        return op(a, b);
    }

    // 广播并逐元素运算
    const [expandedA, expandedB] = broadcast(a, b);
    const result = expandedA.map((row, i) =>
        row.map((val, j) => op(val, expandedB[i][j]))
    );

    // 如果结果是 1×1，返回标量
    if (result.length === 1 && result[0].length === 1) {
        return result[0][0];
    }

    return result;
}

/**
 * Perform element-by-element unary operations on individual arrays
 * @param {*} a - Operand
 * @param {Function} op - Operating function (a) = > result
 * @returns {*} - Results of the operation
 */
export function unaryOp(a, op) {
    if (!Array.isArray(a)) return op(a);

    const arr = to2D(a);
    const result = arr.map(row => row.map(val => op(val)));

    // 如果结果是 1×1，返回标量
    if (result.length === 1 && result[0].length === 1) {
        return result[0][0];
    }

    return result;
}

/**
 * Element-by-element operation on multiple arrays (broadcast support)
 * For scenarios that require multiple sets of operations, such as SUMPRODUCT
 * @param {Array} arrays - Array List
 * @param {Function} op - Operating function (values: Array) = > result
 * @returns {*} - Results of the operation
 */
export function multiElementWise(arrays, op) {
    if (arrays.length === 0) return 0;
    if (arrays.length === 1) {
        const arr = to2D(arrays[0]);
        return arr.map(row => row.map(val => op([val])));
    }

    // 计算最大维度
    let maxRows = 1, maxCols = 1;
    for (const arr of arrays) {
        const [rows, cols] = shape(arr);
        maxRows = Math.max(maxRows, rows);
        maxCols = Math.max(maxCols, cols);
    }

    // 扩展所有数组
    const expanded = arrays.map(arr => expand(arr, maxRows, maxCols));

    // 逐元素运算
    const result = [];
    for (let i = 0; i < maxRows; i++) {
        const row = [];
        for (let j = 0; j < maxCols; j++) {
            const values = expanded.map(arr => arr[i][j]);
            row.push(op(values));
        }
        result.push(row);
    }

    return result;
}

/**
 * Sum of two-dimensional arrays
 * @param {Array<Array>} arr - 2-D Array
 * @returns {number} - Sum
 */
export function sum2D(arr) {
    const arr2D = to2D(arr);
    let total = 0;
    for (const row of arr2D) {
        for (const val of row) {
            const num = Number(val);
            if (!isNaN(num)) total += num;
        }
    }
    return total;
}

/**
 * Convert value to number (Boolean conversion: true→ 1, false→ 0)
 * @param {*} val - Input value
 * @returns {number} - Number
 */
export function toNumber(val) {
    if (typeof val === 'boolean') return val ? 1 : 0;
    if (typeof val === 'number') return val;
    if (val === null || val === undefined || val === '') return 0;
    const num = Number(val);
    if (isNaN(num)) throw new Error('#VALUE!');
    return num;
}

export default {
    shape,
    is2DArray,
    to2D,
    expand,
    broadcast,
    elementWise,
    unaryOp,
    multiElementWise,
    sum2D,
    toNumber
};
