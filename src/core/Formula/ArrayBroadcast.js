/**
 * 数组广播引擎
 * 实现 Excel 风格的数组运算和广播机制
 *
 * 广播规则：
 * - 标量 与 任意数组 → 标量扩展为相同维度
 * - 1×n 行向量 与 m×1 列向量 → 扩展为 m×n 矩阵
 * - 1×n 行向量 与 m×n 矩阵 → 行向量复制 m 次
 * - m×1 列向量 与 m×n 矩阵 → 列向量复制 n 次
 */

/**
 * 获取数组的形状 [rows, cols]
 * @param {*} arr - 输入值（标量、一维数组、二维数组）
 * @returns {[number, number]} - [行数, 列数]
 */
export function shape(arr) {
    if (!Array.isArray(arr)) return [1, 1];  // 标量
    if (arr.length === 0) return [0, 0];
    if (!Array.isArray(arr[0])) return [1, arr.length];  // 一维数组视为行向量
    return [arr.length, arr[0].length];  // 二维数组
}

/**
 * 判断是否为二维数组
 */
export function is2DArray(arr) {
    return Array.isArray(arr) && arr.length > 0 && Array.isArray(arr[0]);
}

/**
 * 将任意值标准化为二维数组
 * @param {*} val - 输入值
 * @returns {Array<Array>} - 二维数组
 */
export function to2D(val) {
    if (!Array.isArray(val)) return [[val]];  // 标量 → 1×1
    if (!Array.isArray(val[0])) return [val];  // 一维 → 1×n 行向量
    return val;  // 已经是二维
}

/**
 * 将二维数组扩展到指定维度
 * @param {Array<Array>} arr - 二维数组
 * @param {number} rows - 目标行数
 * @param {number} cols - 目标列数
 * @returns {Array<Array>} - 扩展后的数组
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
 * 广播两个数组到相同维度
 * @param {*} a - 第一个操作数
 * @param {*} b - 第二个操作数
 * @returns {[Array<Array>, Array<Array>]} - 扩展后的两个数组
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
 * 对两个数组进行逐元素运算（支持广播）
 * @param {*} a - 第一个操作数
 * @param {*} b - 第二个操作数
 * @param {Function} op - 运算函数 (a, b) => result
 * @returns {*} - 运算结果（标量或二维数组）
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
 * 对单个数组进行逐元素一元运算
 * @param {*} a - 操作数
 * @param {Function} op - 运算函数 (a) => result
 * @returns {*} - 运算结果
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
 * 对多个数组进行逐元素运算（支持广播）
 * 用于 SUMPRODUCT 等需要多数组运算的场景
 * @param {Array} arrays - 数组列表
 * @param {Function} op - 运算函数 (values: Array) => result
 * @returns {*} - 运算结果
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
 * 对二维数组求和
 * @param {Array<Array>} arr - 二维数组
 * @returns {number} - 总和
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
 * 将值转换为数字（布尔值转换：true→1, false→0）
 * @param {*} val - 输入值
 * @returns {number} - 数字
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
