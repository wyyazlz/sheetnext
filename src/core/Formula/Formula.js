import { dateTrans, numFmtFun, dateToSerial } from '../Cell/NumFmt.js';
import { elementWise, unaryOp, to2D, shape, expand } from './ArrayBroadcast.js';
import { createMathStatFns } from './FormulaMathStatFns.js';
import { createLookupArrayFns } from './FormulaLookupArrayFns.js';
import { createTextInfoFns } from './FormulaTextInfoFns.js';
import { createDateEngFns } from './FormulaDateEngFns.js';
import {
    collectReferenceTexts,
    hasVolatileFunctionInAst,
    parseFormula,
    splitSheetReference
} from './FormulaAst.js';
import {
    FormulaError,
    coerceLogical,
    coerceNumber,
    coerceText,
    errorCodeOf,
    errorTypeNumber,
    firstErrorValue,
    isBlankValue,
    isErrorCodeString,
    isErrorValue,
    normalizeBlank,
    throwIfError,
    toFormulaError
} from './FormulaValue.js';

const EXCEL_MAX_ROW_INDEX = 1048575;
const EXCEL_MAX_COL_INDEX = 16383;

/**
 * Formula Calculator
 * @title 🔢 Formula
 * @class
 */
export default class Formula {

    /**
     * @param {Object} SN - SheetNext Main Instance
     */
    constructor(SN) {
        this.SN = SN;
        this._lastCalcVolatile = false; // 上次计算的公式是否包含 volatile 函数
        this._lastResultIsDate = false; // 上次计算结果是否为日期类型
    }

    // Excel序列号/Date对象 转 JS Date
    #toDate(val) {
        if (val instanceof Date) return val;
        if (typeof val === 'number') return dateTrans(val);
        return new Date(val);
    }

    /**
     * Whether the last calculation was volatile
     * @type {boolean}
     */
    get lastCalcVolatile() {
        return this._lastCalcVolatile;
    }

    /**
     * Whether the last calculated result is a date type
     * @type {boolean}
     */
    get lastResultIsDate() {
        return this._lastResultIsDate;
    }

    /**
     * Calculation Formula
     * @param {string} formulaStr - Equation String (without Equals)
     * @param {Cell} cell - Current Cell
     * @returns {any}
     */
    calcFormula(formulaStr, cell) {
        try {
            this.currentCell = cell;
            this._lastResultIsDate = false; // 重置日期类型标记
            const ast = parseFormula(formulaStr);
            this._lastCalcVolatile = hasVolatileFunctionInAst(ast);
            return this.#finalizeFormulaResult(this.#evaluateAst(ast));
        } catch (error) {
            if (error?.message?.includes('循环')) this.SN.Utils.toast(error?.message);
            return toFormulaError(error);
        }
    }

    /**
     * Extract all dependent cell coordinates from a formula AST/string.
     * @param {string|Object} formulaOrAst - Formula string without equals or parsed AST
     * @param {string} [defaultSheetName] - Sheet name used by local references
     * @returns {{sheetName:string|null,r:number,c:number}[]}
     */
    parseDeps(formulaOrAst, defaultSheetName = null) {
        const ast = typeof formulaOrAst === 'string' ? parseFormula(formulaOrAst) : formulaOrAst;
        const seen = new Set();
        const deps = [];

        for (const refText of collectReferenceTexts(ast)) {
            for (const dep of this.#depsFromReferenceText(refText, defaultSheetName)) {
                const key = `${dep.sheetName ?? ''}!${dep.r}:${dep.c}`;
                if (seen.has(key)) continue;
                seen.add(key);
                deps.push(dep);
            }
        }

        return deps;
    }

    /**
     * Extract dependency areas without expanding large ranges.
     * @param {string|Object} formulaOrAst - Formula string without equals or parsed AST
     * @param {string} [defaultSheetName] - Sheet name used by local references
     * @returns {{sheetName:string|null,s:{r:number,c:number},e:{r:number,c:number}}[]}
     */
    parseDependencyAreas(formulaOrAst, defaultSheetName = null) {
        const ast = typeof formulaOrAst === 'string' ? parseFormula(formulaOrAst) : formulaOrAst;
        const seen = new Set();
        const areas = [];

        for (const refText of collectReferenceTexts(ast)) {
            const dep = this.#dependencyAreaFromReferenceText(refText, defaultSheetName, true);
            if (!dep) continue;

            const key = `${dep.sheetName ?? ''}!${dep.s.r}:${dep.s.c}:${dep.e.r}:${dep.e.c}`;
            if (seen.has(key)) continue;
            seen.add(key);
            areas.push(dep);
        }

        return areas;
    }

    #evaluateAst(node) {
        switch (node.type) {
            case 'Blank':
                return '';
            case 'Literal':
                return node.value;
            case 'Reference':
                return this.#getCells(node.ref);
            case 'Name':
                return this.#resolveName(node.name);
            case 'ArrayExpression':
                return node.rows.map(row => row.map(item => this.#evaluateAst(item)));
            case 'UnaryExpression':
                return this.#applyUnaryOperator(node.operator, this.#evaluateAst(node.argument));
            case 'PercentExpression':
                return this.#applyPercentOperator(this.#evaluateAst(node.argument));
            case 'ImplicitIntersectionExpression':
                return this.#implicitIntersection(this.#evaluateAst(node.argument));
            case 'SpillReference':
                return this.#resolveSpillReference(node.argument);
            case 'RangeExpression':
                return this.#getCells(this.#rangeTextFromAst(node));
            case 'BinaryExpression':
                return this.#applyBinaryOperator(
                    node.operator,
                    this.#evaluateAst(node.left),
                    this.#evaluateAst(node.right)
                );
            case 'CallExpression':
                return this.#evaluateCallExpression(node);
            default:
                throw new FormulaError('#VALUE!');
        }
    }

    #evaluateCallExpression(node) {
        const name = node.name;
        if (name === 'IF') return this.#evaluateIf(node.args);
        if (name === 'IFS') return this.#evaluateIfs(node.args);
        if (name === 'IFERROR') return this.#evaluateIfError(node.args, false);
        if (name === 'IFNA') return this.#evaluateIfError(node.args, true);
        if (name === 'CHOOSE') return this.#evaluateChoose(node.args);
        if (name === 'SWITCH') return this.#evaluateSwitch(node.args);

        const args = node.args.map(arg => this.#evaluateAstArgument(arg));
        return this.#captureFormulaError(() => this.executeFunction(name, args));
    }

    #evaluateAstArgument(node) {
        return this.#captureFormulaError(() => this.#evaluateAst(node));
    }

    #evaluateIf(args) {
        const condition = this.#captureFormulaError(() => this.#toLogical(this.#gVal(this.#evaluateAst(args[0] ?? { type: 'Blank' }))));
        if (isErrorValue(condition)) return condition;
        if (condition) return args.length > 1 ? this.#evaluateAstArgument(args[1]) : true;
        return args.length > 2 ? this.#evaluateAstArgument(args[2]) : false;
    }

    #evaluateIfs(args) {
        for (let i = 0; i + 1 < args.length; i += 2) {
            const condition = this.#captureFormulaError(() => this.#toLogical(this.#gVal(this.#evaluateAst(args[i]))));
            if (isErrorValue(condition)) return condition;
            if (condition) return this.#evaluateAstArgument(args[i + 1]);
        }
        return toFormulaError('#N/A');
    }

    #evaluateIfError(args, onlyNA) {
        const value = this.#gVal(this.#evaluateAstArgument(args[0] ?? { type: 'Blank' }));
        const err = firstErrorValue(value);
        if (!err) return value;
        if (onlyNA && errorCodeOf(err) !== '#N/A') return value;
        return this.#evaluateAstArgument(args[1] ?? { type: 'Blank' });
    }

    #evaluateChoose(args) {
        const index = this.#captureFormulaError(() => Math.trunc(this.#toNumber(this.#gVal(this.#evaluateAst(args[0] ?? { type: 'Blank' })))));
        if (isErrorValue(index)) return index;
        if (index < 1 || index >= args.length) return toFormulaError('#VALUE!');
        return this.#evaluateAstArgument(args[index]);
    }

    #evaluateSwitch(args) {
        if (!args.length) return toFormulaError('#VALUE!');

        const expression = this.#evaluateAstArgument(args[0]);
        const err = firstErrorValue(expression);
        if (err) return err;

        for (let i = 1; i + 1 < args.length; i += 2) {
            const caseValue = this.#evaluateAstArgument(args[i]);
            const caseErr = firstErrorValue(caseValue);
            if (caseErr) return caseErr;
            if (this.#gVal(caseValue) === this.#gVal(expression)) return this.#evaluateAstArgument(args[i + 1]);
        }

        if (args.length % 2 === 0) return this.#evaluateAstArgument(args[args.length - 1]);
        return toFormulaError('#N/A');
    }

    #applyUnaryOperator(operator, value) {
        const actual = this.#extractValue(value);
        if (operator === '+') return unaryOp(actual, v => this.#captureFormulaError(() => this.#toNumber(v)));
        if (operator === '-') return unaryOp(actual, v => this.#captureFormulaError(() => -this.#toNumber(v)));
        throw new FormulaError('#VALUE!');
    }

    #applyPercentOperator(value) {
        const actual = this.#extractValue(value);
        return unaryOp(actual, v => this.#captureFormulaError(() => this.#toNumber(v) / 100));
    }

    #applyBinaryOperator(operator, left, right) {
        const a = this.#extractValue(left);
        const b = this.#extractValue(right);

        const operations = {
            '+': (x, y) => this.#toNumber(x) + this.#toNumber(y),
            '-': (x, y) => this.#toNumber(x) - this.#toNumber(y),
            '*': (x, y) => this.#toNumber(x) * this.#toNumber(y),
            '/': (x, y) => {
                const bVal = this.#toNumber(y);
                if (bVal === 0) throw new Error('#DIV/0!');
                return this.#toNumber(x) / bVal;
            },
            '^': (x, y) => Math.pow(this.#toNumber(x), this.#toNumber(y)),
            '&': (x, y) => this.#toText(x) + this.#toText(y),
            '=': (x, y) => this.#compareValues(x, y, '='),
            '<': (x, y) => this.#compareValues(x, y, '<'),
            '>': (x, y) => this.#compareValues(x, y, '>'),
            '<=': (x, y) => this.#compareValues(x, y, '<='),
            '>=': (x, y) => this.#compareValues(x, y, '>='),
            '<>': (x, y) => this.#compareValues(x, y, '<>'),
        };

        const fn = operations[operator];
        if (!fn) throw new Error('#NAME?');

        return elementWise(a, b, (x, y) => this.#captureFormulaError(() => {
            const err = firstErrorValue(x, y);
            if (err) throw toFormulaError(err);
            return fn(x, y);
        }));
    }

    #implicitIntersection(value) {
        if (!Array.isArray(value)) return this.#valueFromAny(value);

        const matrix = Array.isArray(value[0]) ? value : [value];
        const currentRow = this.currentCell?.row?.rIndex;
        const currentCol = this.currentCell?.cIndex;

        for (const row of matrix) {
            for (const item of row) {
                if (this.#isCellLike(item) && item.row?.rIndex === currentRow && item.cIndex === currentCol) {
                    return this.#valueFromAny(item);
                }
            }
        }

        if (matrix.length === 1 && currentCol !== undefined) {
            const sameCol = matrix[0].find(item => this.#isCellLike(item) && item.cIndex === currentCol);
            if (sameCol) return this.#valueFromAny(sameCol);
        }

        if (matrix[0]?.length === 1 && currentRow !== undefined) {
            const sameRow = matrix.find(row => this.#isCellLike(row[0]) && row[0].row?.rIndex === currentRow);
            if (sameRow) return this.#valueFromAny(sameRow[0]);
        }

        if (matrix.length === 1 && matrix[0]?.length === 1) return this.#valueFromAny(matrix[0][0]);
        return toFormulaError('#VALUE!');
    }

    #resolveSpillReference(argument) {
        const ref = this.#referenceTextFromAst(argument);
        if (!ref) return toFormulaError('#VALUE!');

        const cells = this.#getCells(ref);
        if (!Array.isArray(cells) || Array.isArray(cells[0]) || cells.length !== 1) return toFormulaError('#VALUE!');

        const sourceCell = cells[0];
        if (sourceCell?.isFormula && !sourceCell.isSpillSource) {
            const _ = sourceCell.calcVal;
        }
        if (!sourceCell || !sourceCell.isSpillSource) return toFormulaError('#REF!');

        const spillArray = sourceCell.spillArray;
        return Array.isArray(spillArray) ? spillArray : toFormulaError('#REF!');
    }

    #finalizeFormulaResult(result) {
        if (Array.isArray(result) && result.length === 1 && !Array.isArray(result[0]) && this.#isCellLike(result[0])) {
            return this.#valueFromAny(result[0]);
        }
        return result;
    }

    #resolveName(name) {
        const definedRef = this.SN._definedNames?.[name] ?? this.SN._definedNames?.[String(name).toUpperCase()];
        if (!definedRef) throw new Error('#NAME?');
        return this.#getCells(definedRef);
    }

    #rangeTextFromAst(node) {
        const left = this.#referenceTextFromAst(node.left);
        const right = this.#referenceTextFromAst(node.right);
        if (!left || !right) throw new FormulaError('#VALUE!');

        const leftRef = splitSheetReference(left);
        const rightRef = splitSheetReference(right);
        const sheetName = rightRef.sheetName ?? leftRef.sheetName;
        return `${sheetName ? this.#quoteSheetName(sheetName) + '!' : ''}${leftRef.address}:${rightRef.address}`;
    }

    #referenceTextFromAst(node) {
        if (node.type === 'Reference') return node.ref;
        if (node.type === 'RangeExpression') return this.#rangeTextFromAst(node);
        return null;
    }

    #quoteSheetName(sheetName) {
        return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)
            ? sheetName
            : `'${String(sheetName).replace(/'/g, "''")}'`;
    }

    #depsFromReferenceText(refText, defaultSheetName) {
        const area = this.#dependencyAreaFromReferenceText(refText, defaultSheetName, false);
        if (!area) return [];
        const deps = [];
        for (let r = area.s.r; r <= area.e.r; r++) {
            for (let c = area.s.c; c <= area.e.c; c++) {
                deps.push({ sheetName: area.sheetName, r, c });
            }
        }
        return deps;
    }

    #dependencyAreaFromReferenceText(refText, defaultSheetName, fullSpan) {
        const ref = splitSheetReference(String(refText ?? ''));
        const sheetName = ref.sheetName ?? defaultSheetName;
        const sheet = sheetName ? this.SN.getSheet?.(sheetName) : this.SN.activeSheet;
        if (!sheet) return null;

        const address = ref.address.replace(/\$/g, '').toUpperCase();
        const area = this.#addressToDependencyArea(address, sheet, fullSpan);
        return { sheetName, s: area.s, e: area.e };
    }

    #addressToDependencyArea(address, sheet, fullSpan) {
        if (fullSpan) {
            const cell = /^([A-Z]{1,3})(\d+)$/i;
            const cellRange = /^([A-Z]{1,3})(\d+):([A-Z]{1,3})(\d+)$/i;
            const colRange = /^([A-Z]{1,3}):([A-Z]{1,3})$/i;
            const rowRange = /^(\d+):(\d+)$/;

            let match = address.match(cell);
            if (match) {
                const r = Number(match[2]) - 1;
                const c = columnNameToIndex(match[1]);
                return { s: { r, c }, e: { r, c } };
            }

            match = address.match(cellRange);
            if (match) {
                const r1 = Number(match[2]) - 1;
                const c1 = columnNameToIndex(match[1]);
                const r2 = Number(match[4]) - 1;
                const c2 = columnNameToIndex(match[3]);
                return {
                    s: { r: Math.min(r1, r2), c: Math.min(c1, c2) },
                    e: { r: Math.max(r1, r2), c: Math.max(c1, c2) }
                };
            }

            match = address.match(colRange);
            if (match) {
                const c1 = columnNameToIndex(match[1]);
                const c2 = columnNameToIndex(match[2]);
                return {
                    s: { r: 0, c: Math.min(c1, c2) },
                    e: { r: EXCEL_MAX_ROW_INDEX, c: Math.max(c1, c2) }
                };
            }

            match = address.match(rowRange);
            if (match) {
                const r1 = Number(match[1]) - 1;
                const r2 = Number(match[2]) - 1;
                return {
                    s: { r: Math.min(r1, r2), c: 0 },
                    e: { r: Math.max(r1, r2), c: EXCEL_MAX_COL_INDEX }
                };
            }
        }

        return sheet.rangeStrToNum(address);
    }

    // 从单元格或值中提取实际值（支持数组）
    #extractValue(val) {
        if (!Array.isArray(val)) return this.#valueFromAny(val);

        // 单元格数组 → 提取 calcVal
        if (val.length > 0 && val[0] && typeof val[0] === 'object' && 'calcVal' in val[0]) {
            // 一维单元格数组
            if (!Array.isArray(val[0])) {
                return val.map(c => this.#valueFromAny(c));
            }
        }

        // 二维单元格数组
        if (val.length > 0 && Array.isArray(val[0])) {
            return val.map(row =>
                row.map(c => this.#valueFromAny(c))
            );
        }

        return val.map(item => this.#valueFromAny(item));
    }

    #captureFormulaError(fn) {
        try {
            return fn();
        } catch (error) {
            return toFormulaError(error);
        }
    }

    #isCellLike(value) {
        return !!(value && typeof value === 'object' && 'calcVal' in value);
    }

    #valueFromAny(value) {
        if (this.#isCellLike(value)) {
            const raw = value.calcVal;
            if ((value.type === 'error' || value._type === 'error') && isErrorCodeString(raw)) {
                return toFormulaError(raw);
            }
            return this.#valueFromAny(raw);
        }
        if (value instanceof Error) return toFormulaError(value);
        if (isErrorValue(value)) return toFormulaError(value);
        return normalizeBlank(value);
    }

    #toNumber(value) {
        return coerceNumber(this.#valueFromAny(value));
    }

    #toLogical(value) {
        return coerceLogical(this.#valueFromAny(value));
    }

    #toText(value) {
        return coerceText(this.#valueFromAny(value));
    }

    #compareValues(left, right, operator) {
        const a = this.#valueFromAny(left);
        const b = this.#valueFromAny(right);
        throwIfError(a);
        throwIfError(b);

        let x = a;
        let y = b;
        if (!isBlankValue(a) && !isBlankValue(b)) {
            const nx = Number(a);
            const ny = Number(b);
            if (Number.isFinite(nx) && Number.isFinite(ny)) {
                x = nx;
                y = ny;
            }
        }

        switch (operator) {
            case '=': return x == y;
            case '<': return x < y;
            case '>': return x > y;
            case '<=': return x <= y;
            case '>=': return x >= y;
            case '<>': return x != y;
            default: throw new FormulaError('#NAME?');
        }
    }

    // 通过地址引用获取单元格
    #getCells(str) {
        const ref = splitSheetReference(String(str).replaceAll('$', ''));
        str = ref.address;
        let sheet = this.SN.activeSheet;
        if (ref.sheetName) sheet = this.SN.getSheet(ref.sheetName);
        if (!sheet) throw new Error('#REF!');

        if (!str.includes(':')) return [sheet.getCell(str)];
        const cells = {};
        sheet.eachCells(str, (r, c) => {
            if (!cells[r]) cells[r] = [];
            cells[r].push(sheet.getCell(r, c));
        });
        return Object.values(cells);
    }

    #flattenValues(value) {
        if (Array.isArray(value)) return value.flatMap(item => this.#flattenValues(item));
        return [this.#valueFromAny(value)];
    }

    #collectNumbers(value, fromReference = false, output = []) {
        if (Array.isArray(value)) {
            value.forEach(item => this.#collectNumbers(item, fromReference, output));
            return output;
        }

        const isRef = fromReference || this.#isCellLike(value);
        const actual = this.#valueFromAny(value);
        throwIfError(actual);

        if (isBlankValue(actual)) return output;
        if (typeof actual === 'number') output.push(actual);
        else if (!isRef && typeof actual === 'boolean') output.push(actual ? 1 : 0);
        else if (!isRef && typeof actual === 'string') {
            const text = actual.trim();
            if (text !== '') {
                const num = Number(text);
                if (Number.isFinite(num)) output.push(num);
            }
        }
        return output;
    }

    #collectLogicals(value, fromReference = false, output = []) {
        if (Array.isArray(value)) {
            value.forEach(item => this.#collectLogicals(item, fromReference, output));
            return output;
        }

        const isRef = fromReference || this.#isCellLike(value);
        const actual = this.#valueFromAny(value);
        throwIfError(actual);

        if (isBlankValue(actual)) return output;
        if (typeof actual === 'boolean') output.push(actual);
        else if (typeof actual === 'number') output.push(actual !== 0);
        else if (!isRef && typeof actual === 'string') output.push(this.#toLogical(actual));
        return output;
    }

    #gVals(stack) {
        return stack.flatMap(x => this.#flattenValues(x));
    }
    #gNums(stack) {
        const nums = [];
        stack.forEach(value => this.#collectNumbers(value, false, nums));
        return nums;
    }
    #gLogicals(stack) {
        const values = [];
        stack.forEach(value => this.#collectLogicals(value, false, values));
        if (values.length === 0) throw new FormulaError('#VALUE!');
        return values;
    }
    #gVal(val) {
        if (Array.isArray(val)) {
            if (val.length === 1 && !Array.isArray(val[0])) {
                return this.#valueFromAny(val[0]);
            } else { // 二维数组
                return val.map(r => Array.isArray(r) ? r.map(c => this.#valueFromAny(c)) : this.#valueFromAny(r))
            }
        }
        return this.#valueFromAny(val)
    }

    //支持返回二维数组，以实现向量化运算
    // gVals:将所有参数/地址引用解构为一维数组（含非数值）
    // gNums:将所有参数/地址引用中数值提取为一维数组（若非数字将使用null占位）
    // gVal:获取单个参数的值，结果是单个值或二维数组

    /**
     * Execute Function
     * @param {string} funcName - Function name
     * @param {Array} stack - Parameter Stack
     * @returns {any}
     */
    executeFunction(funcName, stack) {
        // Helpers for SUBTOTAL and hidden-row aware aggregation
        const pickFirstValue = (val) => {
            if (Array.isArray(val)) {
                if (val.length === 0) return undefined;
                return pickFirstValue(val[0]);
            }
            if (val && typeof val === 'object' && 'calcVal' in val) return val.calcVal;
            return val;
        };

        const getSubtotalMeta = () => {
            const fnValue = pickFirstValue(this.#gVal(stack[0]));
            const fnCode = Number(fnValue);
            if (!Number.isFinite(fnCode)) throw new Error('#VALUE!');

            const funcNum = Math.trunc(fnCode);
            const ignoreManualHidden = funcNum >= 100;
            const baseCode = ignoreManualHidden ? funcNum - 100 : funcNum;
            if (baseCode < 1 || baseCode > 11) throw new Error('#VALUE!');

            return { baseCode, ignoreManualHidden };
        };

        const collectSubtotalValues = (input, ignoreManualHidden, output = []) => {
            if (Array.isArray(input)) {
                input.forEach(item => collectSubtotalValues(item, ignoreManualHidden, output));
                return output;
            }

            if (input && typeof input === 'object' && 'calcVal' in input) {
                const cell = input;
                const row = cell.row;
                const sheet = row?.sheet;
                const isFilterHidden = !!sheet?.AutoFilter?._hiddenRows?.has(row?.rIndex);
                const isRowHidden = !!row?.hidden;

                if (isFilterHidden || (ignoreManualHidden && isRowHidden)) return output;

                const formulaText = typeof cell.editVal === 'string' ? cell.editVal.trim().toUpperCase() : '';
                if (formulaText.startsWith('=SUBTOTAL(')) return output;

                output.push(cell.calcVal);
                return output;
            }

            output.push(input);
            return output;
        };

        const calcSubtotal = (baseCode, values) => {
            const numericValues = values.filter(v => typeof v === 'number' && Number.isFinite(v));

            switch (baseCode) {
                case 1:
                    return numericValues.length ? numericValues.reduce((sum, n) => sum + n, 0) / numericValues.length : 0;
                case 2:
                    return numericValues.length;
                case 3:
                    return values.filter(v => v !== null && v !== undefined && v !== '').length;
                case 4:
                    return numericValues.length ? Math.max(...numericValues) : 0;
                case 5:
                    return numericValues.length ? Math.min(...numericValues) : 0;
                case 6:
                    return numericValues.length ? numericValues.reduce((product, n) => product * n, 1) : 0;
                case 7: {
                    if (numericValues.length <= 1) return 0;
                    const mean = numericValues.reduce((sum, n) => sum + n, 0) / numericValues.length;
                    const variance = numericValues.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / (numericValues.length - 1);
                    return Math.sqrt(variance);
                }
                case 8: {
                    if (!numericValues.length) return 0;
                    const mean = numericValues.reduce((sum, n) => sum + n, 0) / numericValues.length;
                    const variance = numericValues.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / numericValues.length;
                    return Math.sqrt(variance);
                }
                case 9:
                    return numericValues.reduce((sum, n) => sum + n, 0);
                case 10: {
                    if (numericValues.length <= 1) return 0;
                    const mean = numericValues.reduce((sum, n) => sum + n, 0) / numericValues.length;
                    return numericValues.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / (numericValues.length - 1);
                }
                case 11: {
                    if (!numericValues.length) return 0;
                    const mean = numericValues.reduce((sum, n) => sum + n, 0) / numericValues.length;
                    return numericValues.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / numericValues.length;
                }
                default:
                    throw new Error('#VALUE!');
            }
        };

        const isErrorValue = (val) => val instanceof Error || (typeof val === 'string' && val.startsWith('#'));

        const throwWithErrorValue = (val) => {
            if (val instanceof Error) throw val;
            if (typeof val === 'string' && val.startsWith('#')) throw new Error(val);
        };

        const getAggregateMeta = () => {
            const fnValue = pickFirstValue(this.#gVal(stack[0]));
            const optionValue = pickFirstValue(this.#gVal(stack[1]));
            const functionNumRaw = Number(fnValue);
            const optionNumRaw = Number(optionValue);

            if (!Number.isFinite(functionNumRaw) || !Number.isFinite(optionNumRaw)) throw new Error('#VALUE!');

            const functionNum = Math.trunc(functionNumRaw);
            const optionNum = Math.trunc(optionNumRaw);
            if (functionNum < 1 || functionNum > 19 || optionNum < 0 || optionNum > 7) throw new Error('#VALUE!');

            return {
                functionNum,
                ignoreNested: optionNum <= 3,
                ignoreHiddenRows: [1, 3, 5, 7].includes(optionNum),
                ignoreErrors: [2, 3, 6, 7].includes(optionNum),
            };
        };

        const collectAggregateValues = (input, meta, output = []) => {
            if (Array.isArray(input)) {
                input.forEach(item => collectAggregateValues(item, meta, output));
                return output;
            }

            if (input && typeof input === 'object' && 'calcVal' in input) {
                const cell = input;
                const row = cell.row;
                const sheet = row?.sheet;
                const isFilterHidden = !!sheet?.AutoFilter?._hiddenRows?.has(row?.rIndex);
                const isRowHidden = !!row?.hidden;

                if (meta.ignoreHiddenRows && (isFilterHidden || isRowHidden)) return output;

                const formulaText = typeof cell.editVal === 'string' ? cell.editVal.trim().toUpperCase() : '';
                if (meta.ignoreNested && (formulaText.startsWith('=SUBTOTAL(') || formulaText.startsWith('=AGGREGATE('))) return output;

                const cellValue = cell.calcVal;
                if (isErrorValue(cellValue)) {
                    if (meta.ignoreErrors) return output;
                    throwWithErrorValue(cellValue);
                }

                output.push(cellValue);
                return output;
            }

            if (isErrorValue(input)) {
                if (meta.ignoreErrors) return output;
                throwWithErrorValue(input);
            }

            output.push(input);
            return output;
        };

        const getSortedNumbers = (values) => values.filter(v => typeof v === 'number' && Number.isFinite(v)).sort((a, b) => a - b);

        const calcMedian = (sortedNums) => {
            if (!sortedNums.length) return 0;
            const mid = Math.floor(sortedNums.length / 2);
            if (sortedNums.length % 2 === 0) return (sortedNums[mid - 1] + sortedNums[mid]) / 2;
            return sortedNums[mid];
        };

        const calcModeSingle = (sortedNums) => {
            if (!sortedNums.length) throw new Error('#N/A');
            let bestNum = null;
            let bestCount = 0;
            let currentNum = null;
            let currentCount = 0;

            const flush = () => {
                if (currentCount > bestCount || (currentCount === bestCount && bestNum !== null && currentNum < bestNum)) {
                    bestNum = currentNum;
                    bestCount = currentCount;
                }
            };

            sortedNums.forEach(num => {
                if (num === currentNum) {
                    currentCount++;
                } else {
                    if (currentCount > 0) flush();
                    currentNum = num;
                    currentCount = 1;
                }
            });
            if (currentCount > 0) flush();

            if (bestCount <= 1) throw new Error('#N/A');
            return bestNum;
        };

        const percentileInc = (sortedNums, k) => {
            if (!sortedNums.length || !Number.isFinite(k) || k < 0 || k > 1) throw new Error('#NUM!');
            if (sortedNums.length === 1) return sortedNums[0];
            const index = (sortedNums.length - 1) * k;
            const low = Math.floor(index);
            const high = Math.ceil(index);
            if (low === high) return sortedNums[low];
            return sortedNums[low] + (sortedNums[high] - sortedNums[low]) * (index - low);
        };

        const percentileExc = (sortedNums, k) => {
            if (!sortedNums.length || !Number.isFinite(k) || k <= 0 || k >= 1) throw new Error('#NUM!');
            const index = k * (sortedNums.length + 1);
            if (index < 1 || index > sortedNums.length) throw new Error('#NUM!');
            const low = Math.floor(index);
            const high = Math.ceil(index);
            if (low === high) return sortedNums[low - 1];
            const lowVal = sortedNums[low - 1];
            const highVal = sortedNums[high - 1];
            return lowVal + (highVal - lowVal) * (index - low);
        };

        const calcAggregate = (functionNum, values, k) => {
            if (functionNum >= 1 && functionNum <= 11) return calcSubtotal(functionNum, values);

            const sortedNums = getSortedNumbers(values);

            switch (functionNum) {
                case 12:
                    return calcMedian(sortedNums);
                case 13:
                    return calcModeSingle(sortedNums);
                case 14:
                case 15: {
                    const kNum = Number(k);
                    const kth = Math.trunc(kNum);
                    if (!Number.isFinite(kNum) || !Number.isInteger(kNum) || kth < 1 || kth > sortedNums.length) throw new Error('#NUM!');
                    return functionNum === 14 ? sortedNums[sortedNums.length - kth] : sortedNums[kth - 1];
                }
                case 16:
                    return percentileInc(sortedNums, Number(k));
                case 17: {
                    const quartNum = Number(k);
                    const quart = Math.trunc(quartNum);
                    if (!Number.isFinite(quartNum) || !Number.isInteger(quartNum) || quart < 0 || quart > 4) throw new Error('#NUM!');
                    return percentileInc(sortedNums, quart / 4);
                }
                case 18:
                    return percentileExc(sortedNums, Number(k));
                case 19: {
                    const quartNum = Number(k);
                    const quart = Math.trunc(quartNum);
                    if (!Number.isFinite(quartNum) || !Number.isInteger(quartNum) || quart < 1 || quart > 3) throw new Error('#NUM!');
                    return percentileExc(sortedNums, quart / 4);
                }
                default:
                    throw new Error('#VALUE!');
            }
        };

        const getFlatValues = (arg) => {
            const value = this.#gVal(arg);
            if (Array.isArray(value)) return Array.isArray(value[0]) ? value.flat() : [...value];
            return [value];
        };

        const toMatrixValue = (arg) => {
            const value = this.#gVal(arg);
            if (Array.isArray(value)) {
                if (value.length > 0 && Array.isArray(value[0])) return value.map(row => [...row]);
                return [[...value]];
            }
            return [[value]];
        };

        const firstCellFromArg = (arg) => {
            const walk = (input) => {
                if (Array.isArray(input)) {
                    for (const item of input) {
                        const hit = walk(item);
                        if (hit) return hit;
                    }
                    return null;
                }
                if (input && typeof input === 'object' && 'row' in input && 'cIndex' in input) return input;
                return null;
            };
            return walk(arg);
        };

        const escapeRegex = (text) => String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        const buildCriteriaTester = (criterionRaw) => {
            if (typeof criterionRaw === 'string') {
                const criterion = criterionRaw.trim();

                if (/[*?]/.test(criterion)) {
                    const wildcard = '^' + criterion
                        .replace(/([.+^=!:${}()|[\]\\])/g, '\\$1')
                        .replace(/\*/g, '.*')
                        .replace(/\?/g, '.') + '$';
                    const regex = new RegExp(wildcard, 'i');
                    return value => regex.test(String(value ?? ''));
                }

                const match = criterion.match(/^([<>=!]=?|<>)(.*)$/);
                if (match) {
                    let [, op, raw] = match;
                    if (op === '<>') op = '!=';
                    const rawText = raw.trim();
                    const rawNum = Number(rawText);
                    const isNum = rawText !== '' && Number.isFinite(rawNum);

                    return value => {
                        if (isNum) {
                            const v = Number(value);
                            if (!Number.isFinite(v)) return false;
                            switch (op) {
                                case '>': return v > rawNum;
                                case '>=': return v >= rawNum;
                                case '<': return v < rawNum;
                                case '<=': return v <= rawNum;
                                case '=': return v === rawNum;
                                case '!=': return v !== rawNum;
                                default: return false;
                            }
                        }

                        const v = String(value ?? '');
                        switch (op) {
                            case '>': return v > rawText;
                            case '>=': return v >= rawText;
                            case '<': return v < rawText;
                            case '<=': return v <= rawText;
                            case '=': return v.toLowerCase() === rawText.toLowerCase();
                            case '!=': return v.toLowerCase() !== rawText.toLowerCase();
                            default: return false;
                        }
                    };
                }

                return value => String(value ?? '').toLowerCase() === criterion.toLowerCase();
            }

            return value => value == criterionRaw;
        };

        const ensureInteger = (value, errorCode = '#VALUE!') => {
            const num = Number(value);
            if (!Number.isFinite(num) || !Number.isInteger(num)) throw new Error(errorCode);
            return num;
        };

        const getNumericSorted = (argOrValues) => {
            let values;
            if (Array.isArray(argOrValues)) {
                const hasNested = argOrValues.some(item => Array.isArray(item));
                values = hasNested ? argOrValues.flat(Infinity) : argOrValues;
            } else {
                values = getFlatValues(argOrValues);
            }
            return values.map(Number).filter(Number.isFinite).sort((a, b) => a - b);
        };

        const halfWidthText = (text) => String(text)
            .replace(/[\u3000]/g, ' ')
            .replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));

        const fullWidthText = (text) => String(text)
            .replace(/ /g, '\u3000')
            .replace(/[!-~]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0xFEE0));

        const toArrayText = (value) => {
            if (!Array.isArray(value)) return String(value ?? '');
            const matrix = Array.isArray(value[0]) ? value : [value];
            const rows = matrix.map(row =>
                row.map(v => typeof v === 'string' ? `"${v}"` : String(v ?? '')).join(',')
            );
            return `{${rows.join(';')}}`;
        };

        const toDelims = (d) => Array.isArray(d) ? d.map(x => String(x)) : [String(d)];
        const splitByDelims = (input, delims, matchMode = 0) => {
            const pattern = delims.map(escapeRegex).join('|');
            const flags = matchMode === 1 ? 'i' : '';
            return String(input).split(new RegExp(pattern, flags));
        };

        const isoWeekNum = (dateLike) => {
            const date = new Date(dateLike);
            date.setHours(0, 0, 0, 0);
            date.setDate(date.getDate() + 3 - ((date.getDay() + 6) % 7));
            const week1 = new Date(date.getFullYear(), 0, 4);
            return 1 + Math.round(((date.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
        };

        const matchIndex = (arr, lookupValue, matchMode = 0, searchMode = 1) => {
            const indexes = arr.map((_, i) => i);
            const scan = searchMode === -1 ? [...indexes].reverse() : indexes;

            if (matchMode === 2) {
                const pattern = '^' + String(lookupValue)
                    .replace(/([.+^=!:${}()|[\]\\])/g, '\\$1')
                    .replace(/\*/g, '.*')
                    .replace(/\?/g, '.') + '$';
                const regex = new RegExp(pattern, 'i');
                const hit = scan.find(i => regex.test(String(arr[i] ?? '')));
                return hit ?? -1;
            }

            const exact = scan.find(i => arr[i] === lookupValue);
            if (exact !== undefined) return exact;
            if (matchMode === 0) return -1;

            const numLookup = Number(lookupValue);
            const hasNumeric = Number.isFinite(numLookup);
            let bestIndex = -1;

            if (matchMode === -1) {
                let bestVal = -Infinity;
                for (let i = 0; i < arr.length; i++) {
                    const v = hasNumeric ? Number(arr[i]) : arr[i];
                    if ((hasNumeric && Number.isFinite(v) && v <= numLookup && v >= bestVal) ||
                        (!hasNumeric && String(arr[i]) <= String(lookupValue) && (bestIndex === -1 || String(arr[i]) > String(arr[bestIndex])))) {
                        bestVal = v;
                        bestIndex = i;
                    }
                }
            } else if (matchMode === 1) {
                let bestVal = Infinity;
                for (let i = 0; i < arr.length; i++) {
                    const v = hasNumeric ? Number(arr[i]) : arr[i];
                    if ((hasNumeric && Number.isFinite(v) && v >= numLookup && v <= bestVal) ||
                        (!hasNumeric && String(arr[i]) >= String(lookupValue) && (bestIndex === -1 || String(arr[i]) < String(arr[bestIndex])))) {
                        bestVal = v;
                        bestIndex = i;
                    }
                }
            }

            return bestIndex;
        };

        const fnCtx = {
            stack,
            gVal: (arg) => this.#gVal(arg),
            gVals: (args) => this.#gVals(args),
            gNums: (args) => this.#gNums(args),
            toDate: (val) => this.#toDate(val),
            execute: (name, args = stack) => this.executeFunction(name, args),
            setLastResultIsDate: (flag) => { this._lastResultIsDate = !!flag; },
            getCells: (refText) => this.#getCells(refText),
            currentCell: this.currentCell,
            SN: this.SN,
            getFlatValues,
            toMatrixValue,
            firstCellFromArg,
            escapeRegex,
            buildCriteriaTester,
            ensureInteger,
            getNumericSorted,
            halfWidthText,
            fullWidthText,
            toArrayText,
            toDelims,
            splitByDelims,
            isoWeekNum,
            matchIndex,
            percentileInc,
            percentileExc
        };

        const mathStatFns = createMathStatFns(fnCtx);
        const lookupArrayFns = createLookupArrayFns(fnCtx);
        const textInfoFns = createTextInfoFns(fnCtx);
        const dateEngFns = createDateEngFns(fnCtx);

        const F = {
            ...mathStatFns,
            ...lookupArrayFns,
            ...textInfoFns,
            ...dateEngFns,
            PI: () => Math.PI,
            SUM: () => {
                const nums = this.#gNums(stack)
                const r = nums.reduce((total, num) => total + num, 0)
                return r
            },
            SUBTOTAL: () => {
                if (stack.length < 2) throw new Error('#VALUE!');
                const { baseCode, ignoreManualHidden } = getSubtotalMeta();
                const values = collectSubtotalValues(stack.slice(1), ignoreManualHidden);
                return calcSubtotal(baseCode, values);
            },
            AGGREGATE: () => {
                if (stack.length < 3) throw new Error('#VALUE!');
                const meta = getAggregateMeta();

                if (meta.functionNum >= 14) {
                    if (stack.length < 4) throw new Error('#VALUE!');
                    const values = collectAggregateValues(stack[2], meta);
                    const k = pickFirstValue(this.#gVal(stack[3]));
                    return calcAggregate(meta.functionNum, values, k);
                }

                const values = collectAggregateValues(stack.slice(2), meta);
                return calcAggregate(meta.functionNum, values);
            },
            AVERAGE: () => {
                const values = this.#gNums(stack);
                return values.length === 0 ? 0 : values.reduce((total, num) => total + num, 0) / values.length
            },
            MAX: () => Math.max(...this.#gNums(stack)),
            ABS: () => Math.abs(this.#gVal(stack[0])),
            ROUND: () => {
                const number = this.#gVal(stack[0])
                const decimals = this.#gVal(stack[1])
                const factor = Math.pow(10, decimals);
                return Math.round(number * factor) / factor;
            },
            SQRT: () => Math.sqrt(this.#gVal(stack[0])),
            COUNT: () => this.#gNums(stack).length,
            COUNTA: () => this.#gVals(stack).filter(x => x != null && x != '').length,
            VAR: () => {
                const values = this.#gNums(stack);
                const mean = values.reduce((a, b) => a + b) / values.length;
                return values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (values.length - 1);
            },
            STDEV: () => {
                const values = this.#gNums(stack);
                const mean = values.reduce((a, b) => a + b) / values.length;
                return Math.sqrt(values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (values.length - 1));
            },
            IF: () => {
                const condition = this.#toLogical(this.#gVal(stack[0]));
                if (condition) return stack.length > 1 ? this.#gVal(stack[1]) : true;
                return stack.length > 2 ? this.#gVal(stack[2]) : false;
            },
            IFS: () => {
                for (let i = 0; i + 1 < stack.length; i += 2) {
                    if (this.#toLogical(this.#gVal(stack[i]))) return this.#gVal(stack[i + 1]);
                }
                throw new Error('#N/A');
            },
            AND: () => this.#gLogicals(stack).every(value => value),
            NOT: () => !this.#toLogical(this.#gVal(stack[0])),
            OR: () => this.#gLogicals(stack).some(value => value),
            SWITCH: () => {
                const expr = this.#gVal(stack[0]);
                for (let i = 1; i + 1 < stack.length; i += 2) {
                    if (this.#gVal(stack[i]) === expr) return this.#gVal(stack[i + 1]);
                }
                if (stack.length % 2 === 0) return this.#gVal(stack[stack.length - 1]);
                throw new Error('#N/A');
            },
            XOR: () => this.#gLogicals(stack).filter(Boolean).length % 2 === 1,
            CONCATENATE: () => this.#gVals(stack).join(''),
            CONCAT: () => this.#gVals(stack).map(v => String(v ?? '')).join(''),
            LEFT: () => {
                const text = String(this.#gVal(stack[0]));
                const numChars = stack[1] !== undefined ? this.#gVal(stack[1]) : 1;
                if (numChars < 0) return ""; // Excel 行为：负数返回空字符串（或 #VALUE!）
                return text.slice(0, numChars);
            },
            LEFTB: () => this.executeFunction('LEFT', stack),
            RIGHT: () => {
                const text = String(this.#gVal(stack[0]));
                const numChars = stack[1] !== undefined ? this.#gVal(stack[1]) : 1;
                if (numChars < 0) return ""; // Excel 行为：负数返回空字符串（或 #VALUE!）
                if (numChars === 0) return "";
                return text.slice(-numChars);
            },
            RIGHTB: () => this.executeFunction('RIGHT', stack),
            MID: () => {
                const text = String(this.#gVal(stack[0]));
                const startNum = this.#gVal(stack[1]);
                const numChars = this.#gVal(stack[2]);
                // Excel 边界检查：start_num < 1 或 num_chars < 0 返回 #VALUE!
                if (startNum < 1) throw new Error("#VALUE!");
                if (numChars < 0) throw new Error("#VALUE!");
                if (numChars === 0) return "";
                // MID 是 1-indexed，所以需要减 1
                return text.substring(startNum - 1, startNum - 1 + numChars);
            },
            MIDB: () => this.executeFunction('MID', stack),
            LEN: () => String(this.#gVal(stack[0])).length,
            LENB: () => this.executeFunction('LEN', stack),
            LOWER: () => String(this.#gVal(stack[0])).toLowerCase(),
            NOW: () => {
                this._lastResultIsDate = true;
                return dateToSerial(new Date());
            },
            TODAY: () => {
                this._lastResultIsDate = true;
                const date = new Date();
                date.setHours(0, 0, 0, 0);
                return dateToSerial(date);
            },
            YEAR: () => {
                return this.#toDate(this.#gVal(stack[0])).getFullYear()
            },
            MONTH: () => this.#toDate(this.#gVal(stack[0])).getMonth() + 1,
            DAY: () => this.#toDate(this.#gVal(stack[0])).getDate(),
            HLOOKUP: () => {
                const lookupValue = this.#gVal(stack[0]); // 查找值
                const tableArray = this.#gVal(stack[1]);  // 包含数据的二维数组
                const rowIndex = this.#gVal(stack[2]) - 1; // 行索引（从1开始）
                const rangeLookup = this.#gVal(stack[3]) ?? true; // 匹配类型
                if (!Array.isArray(tableArray) || !Array.isArray(tableArray[0])) throw new Error("#REF!");
                const firstRow = tableArray[0];
                let colIndex = -1;
                if (rangeLookup) { // 模糊匹配，找到不大于 lookupValue 的最大值
                    let maxVal = null;
                    colIndex = -1;
                    for (let i = 0; i < firstRow.length; i++) {
                        const cellVal = firstRow[i];
                        if (cellVal <= lookupValue && (maxVal === null || cellVal > maxVal)) {
                            maxVal = cellVal;
                            colIndex = i;
                        }
                    }
                } else {
                    colIndex = firstRow.indexOf(lookupValue); // 精确匹配
                }
                return (colIndex === -1 || rowIndex >= tableArray.length) ? null : tableArray[rowIndex][colIndex];
            },

            // 更多公式实现
            SUMIF: () => {
                const range = this.#gVals(stack[0]);     // 待筛选的值范围
                const criteria = this.#gVal(stack[1]);    // 筛选条件
                const sumRange = stack[2] ? this.#gNums(stack[2]) : this.#gNums(stack[0]); // 求和值的范围


                // 适配向量函数
                const flatCriteria = Array.isArray(criteria) && Array.isArray(criteria[0])
                    ? criteria.flat()
                    : Array.isArray(criteria)
                        ? criteria
                        : null;

                if (flatCriteria) {
                    return flatCriteria.map(crit => {
                        return range.reduce((sum, val, index) => {
                            let match = false;
                            if (typeof crit === 'string') {
                                if (crit.startsWith('>=')) {
                                    match = val >= parseFloat(crit.slice(2));
                                } else if (crit.startsWith('<=')) {
                                    match = val <= parseFloat(crit.slice(2));
                                } else if (crit.startsWith('<>')) {
                                    match = val != parseFloat(crit.slice(2));
                                } else if (crit.startsWith('>')) {
                                    match = val > parseFloat(crit.slice(1));
                                } else if (crit.startsWith('<')) {
                                    match = val < parseFloat(crit.slice(1));
                                } else if (crit.startsWith('=')) {
                                    match = val == parseFloat(crit.slice(1));
                                } else {
                                    match = val == crit;
                                }
                            } else {
                                match = val == crit;
                            }
                            return match ? sum + (sumRange[index] || 0) : sum;
                        }, 0);
                    });
                }

                return range.reduce((sum, val, index) => {
                    // 直接在这里实现条件判断逻辑
                    let match = false;
                    if (typeof criteria === 'string') {
                        if (criteria.startsWith('>=')) {
                            match = val >= parseFloat(criteria.slice(2));
                        } else if (criteria.startsWith('<=')) {
                            match = val <= parseFloat(criteria.slice(2));
                        } else if (criteria.startsWith('<>')) {
                            match = val != parseFloat(criteria.slice(2));
                        } else if (criteria.startsWith('>')) {
                            match = val > parseFloat(criteria.slice(1));
                        } else if (criteria.startsWith('<')) {
                            match = val < parseFloat(criteria.slice(1));
                        } else if (criteria.startsWith('=')) {
                            match = val == parseFloat(criteria.slice(1));
                        } else {
                            // 如果没有运算符，则进行精确匹配
                            match = val == criteria;
                        }
                    } else {
                        // 如果条件不是字符串，则直接比较相等
                        match = val == criteria;
                    }
                    if (match) {
                        return sum + (sumRange[index] || 0);
                    }
                    return sum;
                }, 0);
            },
            LOOKUP: () => {
                const lookupValue = this.#gVal(stack[0]); // 查找值
                let lookupArray = this.#gVal(stack[1]); // 查找数组
                let resultArray = stack.length > 2 ? this.#gVal(stack[2]) : null; // 结果数组

                if (!Array.isArray(lookupArray)) throw new Error("#REF!");

                // 标准化为扁平数组（处理 1×n 行向量和 n×1 列向量）
                let lookupFlat, resultFlat;

                if (!Array.isArray(lookupArray[0])) {
                    // 真正的一维数组
                    lookupFlat = lookupArray;
                } else if (lookupArray.length === 1) {
                    // 1×n 行向量 [[a,b,c]] → [a,b,c]
                    lookupFlat = lookupArray[0];
                } else if (lookupArray[0].length === 1) {
                    // n×1 列向量 [[a],[b],[c]] → [a,b,c]
                    lookupFlat = lookupArray.map(row => row[0]);
                } else {
                    // 真正的二维数组 - 使用第一列查找，返回最后一列
                    const firstCol = lookupArray.map(row => row[0]);
                    let maxVal = null;
                    let resultIndex = -1;
                    for (let i = 0; i < firstCol.length; i++) {
                        const val = firstCol[i];
                        if (val <= lookupValue && (maxVal === null || val > maxVal)) {
                            maxVal = val;
                            resultIndex = i;
                        }
                    }
                    if (resultIndex === -1) return toFormulaError('#N/A');
                    if (resultArray) {
                        const resFlat = Array.isArray(resultArray[0])
                            ? (resultArray.length === 1 ? resultArray[0] : resultArray.map(r => r[resultArray[0].length - 1]))
                            : resultArray;
                        return resFlat[resultIndex];
                    }
                    return lookupArray[resultIndex][lookupArray[0].length - 1];
                }

                // 处理结果数组
                if (resultArray) {
                    if (!Array.isArray(resultArray[0])) {
                        resultFlat = resultArray;
                    } else if (resultArray.length === 1) {
                        resultFlat = resultArray[0];
                    } else if (resultArray[0].length === 1) {
                        resultFlat = resultArray.map(row => row[0]);
                    } else {
                        resultFlat = resultArray[0]; // 取第一行
                    }
                } else {
                    resultFlat = lookupFlat;
                }

                // 在扁平数组中查找（模糊匹配：找不大于 lookupValue 的最大值）
                let maxVal = null;
                let resultIndex = -1;
                for (let i = 0; i < lookupFlat.length; i++) {
                    const val = lookupFlat[i];
                    if (val <= lookupValue && (maxVal === null || val > maxVal)) {
                        maxVal = val;
                        resultIndex = i;
                    }
                }

                return resultIndex === -1 ? toFormulaError('#N/A') : resultFlat[resultIndex];
            },
            XLOOKUP: () => {
                const lookupValue = this.#gVal(stack[0]); // 查找值
                const lookupArray = this.#gVal(stack[1]); // 查找数组
                const returnArray = this.#gVal(stack[2]); // 返回数组
                const notFound = stack.length > 3 ? this.#gVal(stack[3]) : toFormulaError('#N/A'); // 未找到时返回值
                const matchMode = stack.length > 4 ? this.#gVal(stack[4]) : 0; // 匹配模式
                const searchMode = stack.length > 5 ? this.#gVal(stack[5]) : 1; // 搜索模式

                if (!Array.isArray(lookupArray) || !Array.isArray(returnArray)) throw new Error("#REF!");

                let index = -1;
                const arr = lookupArray.flat();

                if (searchMode === 1) { // 从第一个到最后一个
                    switch (matchMode) {
                        case 0: // 精确匹配
                            index = arr.indexOf(lookupValue);
                            break;
                        case -1: // 精确匹配或小于下一个值
                            for (let i = 0; i < arr.length; i++) {
                                if (arr[i] === lookupValue || (arr[i] < lookupValue && (i === arr.length - 1 || arr[i + 1] > lookupValue))) {
                                    index = i;
                                    break;
                                }
                            }
                            break;
                        case 1: // 精确匹配或大于上一个值
                            for (let i = 0; i < arr.length; i++) {
                                if (arr[i] === lookupValue || (arr[i] > lookupValue && (i === 0 || arr[i - 1] < lookupValue))) {
                                    index = i;
                                    break;
                                }
                            }
                            break;
                    }
                } else { // 从最后一个到第一个
                    const reversedArr = [...arr].reverse();
                    switch (matchMode) {
                        case 0:
                            index = arr.lastIndexOf(lookupValue);
                            break;
                        case -1:
                            for (let i = reversedArr.length - 1; i >= 0; i--) {
                                if (reversedArr[i] === lookupValue || (reversedArr[i] < lookupValue && (i === reversedArr.length - 1 || reversedArr[i + 1] > lookupValue))) {
                                    index = arr.length - 1 - i;
                                    break;
                                }
                            }
                            break;
                        case 1:
                            for (let i = reversedArr.length - 1; i >= 0; i--) {
                                if (reversedArr[i] === lookupValue || (reversedArr[i] > lookupValue && (i === 0 || reversedArr[i - 1] < lookupValue))) {
                                    index = arr.length - 1 - i;
                                    break;
                                }
                            }
                            break;
                    }
                }

                return index === -1 ? notFound : returnArray.flat()[index];
            },

            // INDIRECT: 返回文本字符串指定的引用
            INDIRECT: () => {
                const refText = String(this.#gVal(stack[0]));
                // 支持 A1 引用样式，如 "A1", "Sheet1!A1:B5"
                try {
                    return this.#getCells(refText);
                } catch {
                    throw new Error("#REF!");
                }
            },

            // OFFSET: 返回对单元格或范围的引用，该引用与指定单元格有一定的偏移量
            OFFSET: () => {
                const baseRef = stack[0];  // 基准引用（单元格数组）
                const rowOffset = this.#gVal(stack[1]) || 0;
                const colOffset = this.#gVal(stack[2]) || 0;
                const height = stack.length > 3 ? this.#gVal(stack[3]) : null;
                const width = stack.length > 4 ? this.#gVal(stack[4]) : null;

                // 获取基准单元格
                let baseCell;
                if (Array.isArray(baseRef)) {
                    baseCell = Array.isArray(baseRef[0]) ? baseRef[0][0] : baseRef[0];
                } else {
                    throw new Error("#REF!");
                }

                if (!baseCell || typeof baseCell.row === 'undefined') throw new Error("#REF!");

                const sheet = baseCell.row.sheet;
                const baseRow = baseCell.row.rIndex;
                const baseCol = baseCell.cIndex;

                // 计算新位置
                const newRow = baseRow + rowOffset;
                const newCol = baseCol + colOffset;

                if (newRow < 0 || newCol < 0) throw new Error("#REF!");

                // 确定范围大小
                const h = height !== null ? height : 1;
                const w = width !== null ? width : 1;

                if (h <= 0 || w <= 0) throw new Error("#REF!");

                // 返回单个单元格或范围
                if (h === 1 && w === 1) {
                    return [sheet.getCell(newRow, newCol)];
                }

                // 返回二维数组
                const result = [];
                for (let r = 0; r < h; r++) {
                    const row = [];
                    for (let c = 0; c < w; c++) {
                        row.push(sheet.getCell(newRow + r, newCol + c));
                    }
                    result.push(row);
                }
                return result;
            },

            // INDEX: 返回表或区域中的值或值的引用
            INDEX: () => {
                const array = this.#gVal(stack[0]);
                const rowNum = this.#gVal(stack[1]) || 0;
                const colNum = stack.length > 2 ? (this.#gVal(stack[2]) || 0) : 0;

                if (!Array.isArray(array)) {
                    // 单值情况
                    if (rowNum <= 1 && colNum <= 1) return array;
                    throw new Error("#REF!");
                }

                const arr2D = Array.isArray(array[0]) ? array : [array];
                const rows = arr2D.length;
                const cols = arr2D[0].length;

                // rowNum=0 返回整列，colNum=0 返回整行
                if (rowNum === 0 && colNum === 0) {
                    return arr2D;  // 返回整个数组
                }
                if (rowNum === 0) {
                    // 返回指定列
                    if (colNum < 1 || colNum > cols) throw new Error("#REF!");
                    return arr2D.map(row => row[colNum - 1]);
                }
                if (colNum === 0) {
                    // 返回指定行
                    if (rowNum < 1 || rowNum > rows) throw new Error("#REF!");
                    return arr2D[rowNum - 1];
                }

                // 返回单个值
                if (rowNum < 1 || rowNum > rows || colNum < 1 || colNum > cols) {
                    throw new Error("#REF!");
                }
                return arr2D[rowNum - 1][colNum - 1];
            },

            // MATCH: 在范围中搜索指定项，返回该项的相对位置
            MATCH: () => {
                const lookupValue = this.#gVal(stack[0]);
                const lookupArray = this.#gVal(stack[1]);
                const matchType = stack.length > 2 ? this.#gVal(stack[2]) : 1;

                const arr = Array.isArray(lookupArray)
                    ? (Array.isArray(lookupArray[0]) ? lookupArray.flat() : lookupArray)
                    : [lookupArray];

                if (matchType === 0) {
                    // 精确匹配
                    const index = arr.indexOf(lookupValue);
                    return index === -1 ? toFormulaError('#N/A') : index + 1;
                } else if (matchType === 1) {
                    // 小于等于（数组必须升序）
                    let lastMatch = -1;
                    for (let i = 0; i < arr.length; i++) {
                        if (arr[i] <= lookupValue) lastMatch = i;
                        else break;
                    }
                    return lastMatch === -1 ? toFormulaError('#N/A') : lastMatch + 1;
                } else {
                    // 大于等于（数组必须降序）
                    let lastMatch = -1;
                    for (let i = 0; i < arr.length; i++) {
                        if (arr[i] >= lookupValue) lastMatch = i;
                        else break;
                    }
                    return lastMatch === -1 ? toFormulaError('#N/A') : lastMatch + 1;
                }
            },

            // TRANSPOSE: 转置数组
            TRANSPOSE: () => {
                const array = this.#gVal(stack[0]);
                if (!Array.isArray(array)) return [[array]];

                const arr2D = Array.isArray(array[0]) ? array : [array];
                const rows = arr2D.length;
                const cols = arr2D[0].length;

                const result = [];
                for (let c = 0; c < cols; c++) {
                    const row = [];
                    for (let r = 0; r < rows; r++) {
                        row.push(arr2D[r][c]);
                    }
                    result.push(row);
                }
                return result;
            },

            // ==================== 动态数组函数 ====================

            // FILTER: 根据条件筛选数组
            FILTER: () => {
                const array = this.#gVal(stack[0]);
                const include = this.#gVal(stack[1]);
                const ifEmpty = stack.length > 2 ? this.#gVal(stack[2]) : "#CALC!";

                if (!Array.isArray(array)) throw new Error("#VALUE!");

                const arr2D = Array.isArray(array[0]) ? array : [array];
                const inc2D = Array.isArray(include)
                    ? (Array.isArray(include[0]) ? include : [include])
                    : [[include]];

                const result = [];
                const isRowFilter = inc2D.length === arr2D.length && inc2D[0].length === 1;
                const isColFilter = inc2D.length === 1 && inc2D[0].length === arr2D[0].length;

                if (isRowFilter) {
                    // 按行筛选
                    for (let r = 0; r < arr2D.length; r++) {
                        if (inc2D[r][0]) result.push(arr2D[r]);
                    }
                } else if (isColFilter) {
                    // 按列筛选
                    const cols = [];
                    for (let c = 0; c < inc2D[0].length; c++) {
                        if (inc2D[0][c]) cols.push(c);
                    }
                    for (const row of arr2D) {
                        result.push(cols.map(c => row[c]));
                    }
                } else {
                    throw new Error("#VALUE!");
                }

                return result.length > 0 ? result : ifEmpty;
            },

            // UNIQUE: 返回数组中的唯一值
            UNIQUE: () => {
                const array = this.#gVal(stack[0]);
                const byCol = stack.length > 1 ? this.#gVal(stack[1]) : false;
                const exactlyOnce = stack.length > 2 ? this.#gVal(stack[2]) : false;

                if (!Array.isArray(array)) return [[array]];

                const arr2D = Array.isArray(array[0]) ? array : [array];

                if (byCol) {
                    // 按列去重
                    const seen = new Map();
                    const cols = arr2D[0].length;
                    const result = Array(arr2D.length).fill(null).map(() => []);

                    for (let c = 0; c < cols; c++) {
                        const colKey = arr2D.map(r => r[c]).join('\x00');
                        const count = (seen.get(colKey) || 0) + 1;
                        seen.set(colKey, count);
                    }

                    for (let c = 0; c < cols; c++) {
                        const colKey = arr2D.map(r => r[c]).join('\x00');
                        const count = seen.get(colKey);
                        if (!exactlyOnce || count === 1) {
                            if (!exactlyOnce && seen.get(colKey + '_added')) continue;
                            for (let r = 0; r < arr2D.length; r++) {
                                result[r].push(arr2D[r][c]);
                            }
                            seen.set(colKey + '_added', true);
                        }
                    }
                    return result[0].length > 0 ? result : [[""]];
                } else {
                    // 按行去重
                    const seen = new Map();
                    for (const row of arr2D) {
                        const rowKey = row.join('\x00');
                        seen.set(rowKey, (seen.get(rowKey) || 0) + 1);
                    }

                    const result = [];
                    const added = new Set();
                    for (const row of arr2D) {
                        const rowKey = row.join('\x00');
                        const count = seen.get(rowKey);
                        if (!exactlyOnce || count === 1) {
                            if (!exactlyOnce && added.has(rowKey)) continue;
                            result.push([...row]);
                            added.add(rowKey);
                        }
                    }
                    return result.length > 0 ? result : [[""]];
                }
            },

            // SORT: 对数组排序
            SORT: () => {
                const array = this.#gVal(stack[0]);
                const sortIndex = stack.length > 1 ? this.#gVal(stack[1]) : 1;
                const sortOrder = stack.length > 2 ? this.#gVal(stack[2]) : 1;  // 1=升序, -1=降序
                const byCol = stack.length > 3 ? this.#gVal(stack[3]) : false;

                if (!Array.isArray(array)) return [[array]];

                const arr2D = Array.isArray(array[0]) ? array : [array];
                const result = arr2D.map(row => [...row]);

                if (byCol) {
                    // 按列排序（转置 → 排序 → 转置回来）
                    const transposed = result[0].map((_, c) => result.map(r => r[c]));
                    transposed.sort((a, b) => {
                        const aVal = a[sortIndex - 1];
                        const bVal = b[sortIndex - 1];
                        if (aVal < bVal) return -sortOrder;
                        if (aVal > bVal) return sortOrder;
                        return 0;
                    });
                    return transposed[0].map((_, r) => transposed.map(c => c[r]));
                } else {
                    // 按行排序
                    result.sort((a, b) => {
                        const aVal = a[sortIndex - 1];
                        const bVal = b[sortIndex - 1];
                        if (aVal < bVal) return -sortOrder;
                        if (aVal > bVal) return sortOrder;
                        return 0;
                    });
                    return result;
                }
            },

            // SORTBY: 根据另一个数组或范围对数组排序
            SORTBY: () => {
                const array = this.#gVal(stack[0]);
                if (!Array.isArray(array)) return [[array]];

                const arr2D = Array.isArray(array[0]) ? array : [array];
                const result = arr2D.map((row, i) => ({ row: [...row], index: i }));

                // 处理多个排序键
                for (let i = 1; i < stack.length; i += 2) {
                    const byArray = this.#gVal(stack[i]);
                    const order = i + 1 < stack.length ? this.#gVal(stack[i + 1]) : 1;

                    const byArr = Array.isArray(byArray)
                        ? (Array.isArray(byArray[0]) ? byArray.map(r => r[0]) : byArray)
                        : [byArray];

                    result.sort((a, b) => {
                        const aVal = byArr[a.index];
                        const bVal = byArr[b.index];
                        if (aVal < bVal) return -order;
                        if (aVal > bVal) return order;
                        return 0;
                    });
                }

                return result.map(item => item.row);
            },

            // SEQUENCE: 生成数字序列
            SEQUENCE: () => {
                const rows = this.#gVal(stack[0]);
                const cols = stack.length > 1 ? this.#gVal(stack[1]) : 1;
                const start = stack.length > 2 ? this.#gVal(stack[2]) : 1;
                const step = stack.length > 3 ? this.#gVal(stack[3]) : 1;

                if (rows <= 0 || cols <= 0) throw new Error("#VALUE!");

                const result = [];
                let value = start;
                for (let r = 0; r < rows; r++) {
                    const row = [];
                    for (let c = 0; c < cols; c++) {
                        row.push(value);
                        value += step;
                    }
                    result.push(row);
                }
                return result;
            },

            // RANDARRAY: 生成随机数数组
            RANDARRAY: () => {
                const rows = stack.length > 0 ? this.#gVal(stack[0]) : 1;
                const cols = stack.length > 1 ? this.#gVal(stack[1]) : 1;
                const min = stack.length > 2 ? this.#gVal(stack[2]) : 0;
                const max = stack.length > 3 ? this.#gVal(stack[3]) : 1;
                const wholeNumber = stack.length > 4 ? this.#gVal(stack[4]) : false;

                if (rows <= 0 || cols <= 0) throw new Error("#VALUE!");

                const result = [];
                for (let r = 0; r < rows; r++) {
                    const row = [];
                    for (let c = 0; c < cols; c++) {
                        let val = Math.random() * (max - min) + min;
                        if (wholeNumber) val = Math.floor(val);
                        row.push(val);
                    }
                    result.push(row);
                }
                return result;
            },
            // CHOOSECOLS: 从数组中选择指定列
            CHOOSECOLS: () => {
                const array = this.#gVal(stack[0]);
                if (!Array.isArray(array)) throw new Error("#VALUE!");

                const arr2D = Array.isArray(array[0]) ? array : [array];
                const colIndices = stack.slice(1).map(s => this.#gVal(s));

                const result = [];
                for (const row of arr2D) {
                    const newRow = [];
                    for (const colIdx of colIndices) {
                        const idx = colIdx > 0 ? colIdx - 1 : row.length + colIdx;
                        if (idx < 0 || idx >= row.length) throw new Error("#VALUE!");
                        newRow.push(row[idx]);
                    }
                    result.push(newRow);
                }
                return result;
            },

            // CHOOSEROWS: 从数组中选择指定行
            CHOOSEROWS: () => {
                const array = this.#gVal(stack[0]);
                if (!Array.isArray(array)) throw new Error("#VALUE!");

                const arr2D = Array.isArray(array[0]) ? array : [array];
                const rowIndices = stack.slice(1).map(s => this.#gVal(s));

                const result = [];
                for (const rowIdx of rowIndices) {
                    const idx = rowIdx > 0 ? rowIdx - 1 : arr2D.length + rowIdx;
                    if (idx < 0 || idx >= arr2D.length) throw new Error("#VALUE!");
                    result.push([...arr2D[idx]]);
                }
                return result;
            },

            // VSTACK: 垂直堆叠数组
            VSTACK: () => {
                const result = [];
                for (const arg of stack) {
                    const arr = this.#gVal(arg);
                    const arr2D = Array.isArray(arr)
                        ? (Array.isArray(arr[0]) ? arr : [arr])
                        : [[arr]];
                    result.push(...arr2D);
                }
                return result;
            },

            // HSTACK: 水平堆叠数组
            HSTACK: () => {
                const arrays = stack.map(arg => {
                    const arr = this.#gVal(arg);
                    return Array.isArray(arr)
                        ? (Array.isArray(arr[0]) ? arr : [arr])
                        : [[arr]];
                });

                // 找到最大行数
                const maxRows = Math.max(...arrays.map(a => a.length));

                const result = [];
                for (let r = 0; r < maxRows; r++) {
                    const row = [];
                    for (const arr of arrays) {
                        const srcRow = arr[r] || arr[arr.length - 1] || [];
                        row.push(...srcRow);
                    }
                    result.push(row);
                }
                return result;
            },

            // WRAPCOLS: 按列包装数组
            WRAPCOLS: () => {
                const vector = this.#gVal(stack[0]);
                const wrapCount = this.#gVal(stack[1]);
                const padWith = stack.length > 2 ? this.#gVal(stack[2]) : toFormulaError('#N/A');

                const arr = Array.isArray(vector) ? vector.flat() : [vector];
                const result = [];
                const cols = Math.ceil(arr.length / wrapCount);

                for (let r = 0; r < wrapCount; r++) {
                    const row = [];
                    for (let c = 0; c < cols; c++) {
                        const idx = c * wrapCount + r;
                        row.push(idx < arr.length ? arr[idx] : padWith);
                    }
                    result.push(row);
                }
                return result;
            },

            // WRAPROWS: 按行包装数组
            WRAPROWS: () => {
                const vector = this.#gVal(stack[0]);
                const wrapCount = this.#gVal(stack[1]);
                const padWith = stack.length > 2 ? this.#gVal(stack[2]) : toFormulaError('#N/A');

                const arr = Array.isArray(vector) ? vector.flat() : [vector];
                const result = [];

                for (let i = 0; i < arr.length; i += wrapCount) {
                    const row = arr.slice(i, i + wrapCount);
                    while (row.length < wrapCount) row.push(padWith);
                    result.push(row);
                }
                return result;
            },

            // TOCOL: 将数组转换为单列
            TOCOL: () => {
                const array = this.#gVal(stack[0]);
                const ignore = stack.length > 1 ? this.#gVal(stack[1]) : 0;
                const scanByCol = stack.length > 2 ? this.#gVal(stack[2]) : false;

                const arr2D = Array.isArray(array)
                    ? (Array.isArray(array[0]) ? array : [array])
                    : [[array]];

                const result = [];
                if (scanByCol) {
                    for (let c = 0; c < arr2D[0].length; c++) {
                        for (let r = 0; r < arr2D.length; r++) {
                            const val = arr2D[r][c];
                            if (ignore === 1 && val === "") continue;
                            if (ignore === 2 && (val instanceof Error || String(val).startsWith('#'))) continue;
                            if (ignore === 3 && (val === "" || val instanceof Error || String(val).startsWith('#'))) continue;
                            result.push([val]);
                        }
                    }
                } else {
                    for (const row of arr2D) {
                        for (const val of row) {
                            if (ignore === 1 && val === "") continue;
                            if (ignore === 2 && (val instanceof Error || String(val).startsWith('#'))) continue;
                            if (ignore === 3 && (val === "" || val instanceof Error || String(val).startsWith('#'))) continue;
                            result.push([val]);
                        }
                    }
                }
                return result.length > 0 ? result : [[""]];
            },

            // TOROW: 将数组转换为单行
            TOROW: () => {
                const array = this.#gVal(stack[0]);
                const ignore = stack.length > 1 ? this.#gVal(stack[1]) : 0;
                const scanByCol = stack.length > 2 ? this.#gVal(stack[2]) : false;

                const arr2D = Array.isArray(array)
                    ? (Array.isArray(array[0]) ? array : [array])
                    : [[array]];

                const result = [];
                if (scanByCol) {
                    for (let c = 0; c < arr2D[0].length; c++) {
                        for (let r = 0; r < arr2D.length; r++) {
                            const val = arr2D[r][c];
                            if (ignore === 1 && val === "") continue;
                            if (ignore === 2 && (val instanceof Error || String(val).startsWith('#'))) continue;
                            if (ignore === 3 && (val === "" || val instanceof Error || String(val).startsWith('#'))) continue;
                            result.push(val);
                        }
                    }
                } else {
                    for (const row of arr2D) {
                        for (const val of row) {
                            if (ignore === 1 && val === "") continue;
                            if (ignore === 2 && (val instanceof Error || String(val).startsWith('#'))) continue;
                            if (ignore === 3 && (val === "" || val instanceof Error || String(val).startsWith('#'))) continue;
                            result.push(val);
                        }
                    }
                }
                return [result.length > 0 ? result : [""]];
            },

            VLOOKUP: () => {
                const lookupValue = this.#gVal(stack[0]); // 查找值
                const tableArray = this.#gVal(stack[1]);  // 包含数据的二维数组
                const colIndex = this.#gVal(stack[2]) - 1; // 列索引（从1开始）
                const rangeLookup = this.#gVal(stack[3]) ?? true; // 匹配类型

                if (!Array.isArray(tableArray) || !Array.isArray(tableArray[0])) throw new Error("#REF!");
                if (colIndex >= tableArray[0].length) throw new Error("#REF!");

                const firstCol = tableArray.map(row => row[0]);
                let rowIndex = -1;

                if (rangeLookup) { // 模糊匹配，找到不大于 lookupValue 的最大值
                    let maxVal = null;
                    for (let i = 0; i < firstCol.length; i++) {
                        const val = firstCol[i];
                        if (val <= lookupValue && (maxVal === null || val > maxVal)) {
                            maxVal = val;
                            rowIndex = i;
                        }
                    }
                } else { // 精确匹配
                    rowIndex = firstCol.indexOf(lookupValue);
                }

                return rowIndex === -1 ? toFormulaError('#N/A') : tableArray[rowIndex][colIndex];
            },
            TRIM: () => String(this.#gVal(stack[0])).replace(/^\s+|\s+$/g, '').replace(/\s+/g, ' '),
            MIN: () => Math.min(...this.#gNums(stack)),
            POWER: () => {
                const base = this.#gVal(stack[0])
                const exponent = this.#gVal(stack[1])
                return Math.pow(base, exponent);
            },
            ROW: () => {
                if (stack[0]?.[0]) return stack[0][0]?.row.rIndex + 1
                return this.currentCell.row.rIndex + 1
            },
            COLUMN: () => {
                if (stack[0]?.[0]) return stack[0][0]?.cIndex + 1
                return this.currentCell.cIndex + 1
            },
            ADDRESS: () => {
                const rowNum = Number(this.#gVal(stack[0]));
                const colNum = Number(this.#gVal(stack[1]));
                const absNumRaw = stack[2] !== undefined ? Number(this.#gVal(stack[2])) : 1;
                const a1Raw = stack[3] !== undefined ? this.#gVal(stack[3]) : true;
                const sheetTextRaw = stack[4] !== undefined ? this.#gVal(stack[4]) : null;

                if (!Number.isFinite(rowNum) || !Number.isFinite(colNum)) throw new Error('#VALUE!');
                const row = Math.trunc(rowNum);
                const col = Math.trunc(colNum);
                if (row < 1 || col < 1) throw new Error('#VALUE!');

                const absNum = Number.isFinite(absNumRaw) ? Math.trunc(absNumRaw) : 1;
                if (![1, 2, 3, 4].includes(absNum)) throw new Error('#VALUE!');

                const a1 = !(a1Raw === false || a1Raw === 0 || String(a1Raw).toUpperCase() === 'FALSE');
                const colText = this.SN.Utils.numToChar(col - 1);

                let ref;
                if (a1) {
                    const isColAbs = absNum === 1 || absNum === 3;
                    const isRowAbs = absNum === 1 || absNum === 2;
                    ref = `${isColAbs ? '$' : ''}${colText}${isRowAbs ? '$' : ''}${row}`;
                } else {
                    const rowRef = (absNum === 1 || absNum === 2) ? `R${row}` : `R[${row}]`;
                    const colRef = (absNum === 1 || absNum === 3) ? `C${col}` : `C[${col}]`;
                    ref = `${rowRef}${colRef}`;
                }

                if (sheetTextRaw === null || sheetTextRaw === undefined || sheetTextRaw === '') return ref;

                const sheetText = String(sheetTextRaw);
                const escapedSheet = sheetText.replace(/'/g, "''");
                const needQuote = /[^A-Za-z0-9_.]/.test(sheetText);
                return `${needQuote ? `'${escapedSheet}'` : escapedSheet}!${ref}`;
            },
            AREAS: () => {
                if (!stack.length) throw new Error('#VALUE!');
                return stack.length;
            },
            UPPER: () => String(this.#gVal(stack[0])).toUpperCase(),
            OR: () => this.#gLogicals(stack).some(value => value),
            ACOS: () => Math.acos(this.#gVal(stack[0])),
            ASIN: () => Math.asin(this.#gVal(stack[0])),
            ATAN: () => Math.atan(this.#gVal(stack[0])),
            ACOT: () => {
                const num = Number(this.#gVal(stack[0]));
                if (!Number.isFinite(num)) throw new Error('#VALUE!');
                if (num === 0) return Math.PI / 2;
                return Math.atan(1 / num) + (num < 0 ? Math.PI : 0);
            },
            ATAN2: () => Math.atan2(this.#gVal(stack[0]), this.#gVal(stack[1])),
            COS: () => Math.cos(this.#gVal(stack[0])),
            CEILING: () => {
                const number = this.#gVal(stack[0]);
                const significance = this.#gVal(stack[1]);
                return Math.ceil(number / significance) * significance;
            },
            ODD: () => {
                const number = Math.ceil(this.#gVal(stack[0]));
                return number % 2 === 0 ? number + 1 : number;
            },
            EVEN: () => {
                const number = Math.ceil(this.#gVal(stack[0]));
                return number % 2 === 0 ? number : number + 1;
            },
            FLOOR: () => {
                const number = this.#gVal(stack[0]);
                const significance = this.#gVal(stack[1]);
                return Math.floor(number / significance) * significance;
            },
            LN: () => Math.log(this.#gVal(stack[0])),
            SIN: () => Math.sin(this.#gVal(stack[0])),
            TAN: () => Math.tan(this.#gVal(stack[0])),
            SIGN: () => {
                const number = this.#gVal(stack[0]);
                return number > 0 ? 1 : number < 0 ? -1 : 0;
            },
            GCD: () => {
                const values = this.#gNums(stack);
                const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
                return values.reduce((acc, val) => gcd(acc, val));
            },
            LCM: () => {
                const values = this.#gNums(stack);
                const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
                const lcm = (a, b) => Math.abs(a * b) / gcd(a, b);
                return values.reduce((acc, val) => lcm(acc, val));
            },
            PRODUCT: () => this.#gNums(stack).reduce((acc, num) => acc * num, 1),
            MOD: () => this.#gVal(stack[0]) % this.#gVal(stack[1]),
            QUOTIENT: () => Math.floor(this.#gVal(stack[0]) / this.#gVal(stack[1])),
            INT: () => Math.floor(this.#gVal(stack[0])),
            MROUND: () => {
                const number = this.#gVal(stack[0]);
                const multiple = this.#gVal(stack[1]);
                return Math.round(number / multiple) * multiple;
            },
            ROUNDDOWN: () => {
                const number = this.#gVal(stack[0]);
                const decimals = this.#gVal(stack[1]);
                const factor = Math.pow(10, decimals);
                return Math.floor(number * factor) / factor;
            },
            ROUNDUP: () => {
                const number = this.#gVal(stack[0]);
                const decimals = this.#gVal(stack[1]);
                const factor = Math.pow(10, decimals);
                return Math.ceil(number * factor) / factor;
            },
            TRUNC: () => {
                const number = this.#gVal(stack[0]);
                const decimals = this.#gVal(stack[1]) ?? 0;
                const factor = Math.pow(10, decimals);
                return Math.trunc(number * factor) / factor;
            },
            EXP: () => Math.exp(this.#gVal(stack[0])),
            LOG: () => {
                const number = this.#gVal(stack[0]);
                const base = this.#gVal(stack[1]) ?? 10;
                return Math.log(number) / Math.log(base);
            },
            LOG10: () => Math.log10(this.#gVal(stack[0])),
            SUMPRODUCT: () => {
                // 获取所有参数并转为二维数组
                let arrays = stack.map(arg => {
                    const val = this.#gVal(arg);
                    return to2D(val);
                });

                if (arrays.length === 0) return 0;

                // 计算最大维度
                let maxRows = 1, maxCols = 1;
                for (const arr of arrays) {
                    const [rows, cols] = shape(arr);
                    maxRows = Math.max(maxRows, rows);
                    maxCols = Math.max(maxCols, cols);
                }

                // 广播所有数组到相同维度
                arrays = arrays.map(arr => expand(arr, maxRows, maxCols));

                // 逐元素相乘后求和
                let sum = 0;
                for (let i = 0; i < maxRows; i++) {
                    for (let j = 0; j < maxCols; j++) {
                        let product = 1;
                        for (const arr of arrays) {
                            const val = arr[i][j];
                            // true → 1, false → 0, 数字保持, 其他转数字
                            product *= this.#toNumber(val);
                        }
                        sum += product;
                    }
                }
                return sum;
            },
            SUMSQ: () => this.#gNums(stack).reduce((total, num) => total + Math.pow(num, 2), 0),
            SUMX2MY2: () => {
                if (stack.length < 2) throw new Error('#VALUE!');
                const arrayX = this.#gNums([stack[0]]).filter(v => typeof v === 'number' && Number.isFinite(v));
                const arrayY = this.#gNums([stack[1]]).filter(v => typeof v === 'number' && Number.isFinite(v));
                if (arrayX.length !== arrayY.length) throw new Error('#VALUE!');
                return arrayX.reduce((total, x, i) => total + Math.pow(x, 2) - Math.pow(arrayY[i], 2), 0);
            },
            SUMX2PY2: () => {
                if (stack.length < 2) throw new Error('#VALUE!');
                const arrayX = this.#gNums([stack[0]]).filter(v => typeof v === 'number' && Number.isFinite(v));
                const arrayY = this.#gNums([stack[1]]).filter(v => typeof v === 'number' && Number.isFinite(v));
                if (arrayX.length !== arrayY.length) throw new Error('#VALUE!');
                return arrayX.reduce((total, x, i) => total + Math.pow(x, 2) + Math.pow(arrayY[i], 2), 0);
            },
            SUMXMY2: () => {
                if (stack.length < 2) throw new Error('#VALUE!');
                const arrayX = this.#gNums([stack[0]]).filter(v => typeof v === 'number' && Number.isFinite(v));
                const arrayY = this.#gNums([stack[1]]).filter(v => typeof v === 'number' && Number.isFinite(v));
                if (arrayX.length !== arrayY.length) throw new Error('#VALUE!');
                return arrayX.reduce((total, x, i) => total + Math.pow(x - arrayY[i], 2), 0);
            },
            SERIESSUM: () => {
                if (stack.length < 4) throw new Error('#VALUE!');
                const x = this.#gVal(stack[0]);
                const n = this.#gVal(stack[1]);
                const m = this.#gVal(stack[2]);
                const coefficients = this.#gVals([stack[3]])
                    .map(v => Number(v))
                    .filter(v => Number.isFinite(v));
                if (!coefficients.length) throw new Error('#VALUE!');

                return coefficients.reduce((sum, coef, i) => sum + coef * Math.pow(x, n + i * m), 0);
            },
            SQRTPI: () => Math.sqrt(Math.PI * this.#gVal(stack[0])),
            DEGREES: () => this.#gVal(stack[0]) * (180 / Math.PI),
            RADIANS: () => this.#gVal(stack[0]) * (Math.PI / 180),
            COSH: () => Math.cosh(this.#gVal(stack[0])),
            ACOSH: () => Math.acosh(this.#gVal(stack[0])),
            SINH: () => Math.sinh(this.#gVal(stack[0])),
            ASINH: () => Math.asinh(this.#gVal(stack[0])),
            TANH: () => Math.tanh(this.#gVal(stack[0])),
            ATANH: () => Math.atanh(this.#gVal(stack[0])),
            ACOTH: () => {
                const num = Number(this.#gVal(stack[0]));
                if (!Number.isFinite(num)) throw new Error('#VALUE!');
                if (Math.abs(num) <= 1) throw new Error('#NUM!');
                return 0.5 * Math.log((num + 1) / (num - 1));
            },
            MDETERM: () => {
                const matrix = this.#gVal(stack[0]);
                if (!Array.isArray(matrix) || matrix.length !== matrix[0].length) throw new Error("#VALUE!");
                const det = (m) => m.length === 1 ? m[0][0] : m.reduce((sum, val, i) => sum + (i % 2 ? -1 : 1) * val[0] * det(m.slice(1).map(row => row.slice(1))), 0);
                return det(matrix);
            },
            MINVERSE: () => {
                const matrix = this.#gVal(stack[0]);
                if (!Array.isArray(matrix) || matrix.length !== matrix[0].length) throw new Error("#VALUE!");
                const n = matrix.length;
                const identity = Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => i === j ? 1 : 0));
                const augmented = matrix.map((row, i) => [...row, ...identity[i]]);
                for (let i = 0; i < n; i++) {
                    let factor = augmented[i][i];
                    if (factor === 0) throw new Error("#DIV/0!");
                    for (let j = 0; j < n * 2; j++) augmented[i][j] /= factor;
                    for (let k = 0; k < n; k++) {
                        if (k !== i) {
                            factor = augmented[k][i];
                            for (let j = 0; j < n * 2; j++) augmented[k][j] -= factor * augmented[i][j];
                        }
                    }
                }
                return augmented.map(row => row.slice(n));
            },
            MMULT: () => {
                const matrixA = this.#gVal(stack[0]);
                const matrixB = this.#gVal(stack[1]);
                if (!Array.isArray(matrixA) || !Array.isArray(matrixB) || matrixA[0].length !== matrixB.length) throw new Error("#VALUE!");
                return matrixA.map(row => matrixB[0].map((_, colIndex) => row.reduce((sum, cell, rowIndex) => sum + cell * matrixB[rowIndex][colIndex], 0)));
            },
            FACT: () => {
                const num = this.#gVal(stack[0]);
                if (num < 0 || !Number.isInteger(num)) throw new Error("#NUM!");
                return num === 0 ? 1 : Array.from({ length: num }, (_, i) => i + 1).reduce((acc, n) => acc * n, 1);
            },
            FACTDOUBLE: () => {
                const n = this.#gVal(stack[0]);
                if (n < 0 || !Number.isInteger(n)) throw new Error("#NUM!");
                let result = 1;
                for (let i = n; i > 0; i -= 2) result *= i;
                return result;
            },

            RAND: () => Math.random(),

            RANDBETWEEN: () => {
                const min = this.#gVal(stack[0]);
                const max = this.#gVal(stack[1]);
                return Math.floor(Math.random() * (max - min + 1)) + min;
            },

            ROMAN: () => {
                const num = this.#gVal(stack[0]);
                if (num <= 0 || num > 3999 || !Number.isInteger(num)) throw new Error("#VALUE!");
                const romanNumeralMap = [
                    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"],
                    [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"],
                    [5, "V"], [4, "IV"], [1, "I"]
                ];
                let roman = "";
                let remaining = num;
                for (const [value, symbol] of romanNumeralMap) {
                    while (remaining >= value) {
                        roman += symbol;
                        remaining -= value;
                    }
                }
                return roman;
            },
            ARABIC: () => {
                const roman = String(this.#gVal(stack[0]) ?? '').toUpperCase().replace(/\s+/g, '');
                if (!roman) throw new Error('#VALUE!');

                const validRoman = /^M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$/;
                if (!validRoman.test(roman)) throw new Error('#VALUE!');

                const romanMap = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
                let total = 0;
                for (let i = 0; i < roman.length; i++) {
                    const current = romanMap[roman[i]];
                    const next = romanMap[roman[i + 1]] || 0;
                    total += current < next ? -current : current;
                }
                return total;
            },

            CEILING_PRECISE: () => {
                const number = this.#gVal(stack[0]);
                const significance = Math.abs(this.#gVal(stack[1]) || 1);
                return Math.ceil(number / significance) * significance;
            },

            ISO_CEILING: () => {
                const number = this.#gVal(stack[0]);
                const significance = Math.abs(this.#gVal(stack[1]) || 1);
                return Math.ceil(number / significance) * significance;
            },

            FLOOR_PRECISE: () => {
                const number = this.#gVal(stack[0]);
                const significance = Math.abs(this.#gVal(stack[1]) || 1);
                return Math.floor(number / significance) * significance;
            },

            MUNIT: () => {
                const size = this.#gVal(stack[0]);
                if (size <= 0 || !Number.isInteger(size)) throw new Error("#VALUE!");
                const matrix = Array(size).fill(null).map((_, i) =>
                    Array(size).fill(0).map((_, j) => (i === j ? 1 : 0))
                );
                return matrix;
            },
            IFERROR: () => {
                const value = this.#gVal(stack[0]);
                return firstErrorValue(value) ? this.#gVal(stack[1]) : value;
            },
            IFNA: () => {
                const value = this.#gVal(stack[0]);
                const err = firstErrorValue(value);
                return err && errorCodeOf(err) === '#N/A' ? this.#gVal(stack[1]) : value;
            },
            TRUE: () => true,
            FALSE: () => false,
            DATE: () => {
                this._lastResultIsDate = true;
                const year = this.#gVal(stack[0]);
                const month = this.#gVal(stack[1]) - 1; // JavaScript月份从0开始
                const day = this.#gVal(stack[2]);
                return dateToSerial(new Date(year, month, day));
            },
            TIME: () => {
                this._lastResultIsDate = true;
                const hours = this.#gVal(stack[0]);
                const minutes = this.#gVal(stack[1]);
                const seconds = this.#gVal(stack[2]);
                // TIME 返回一天内的小数部分
                return (hours * 3600 + minutes * 60 + seconds) / 86400;
            },
            DATEVALUE: () => {
                this._lastResultIsDate = true;
                const dateString = this.#gVal(stack[0]);
                const date = new Date(dateString);
                date.setHours(0, 0, 0, 0);
                return dateToSerial(date);
            },
            TIMEVALUE: () => {
                const timeString = this.#gVal(stack[0]);
                const [hours, minutes, seconds] = timeString.split(':').map(Number);
                return hours * 3600 + minutes * 60 + (seconds || 0);
            },
            HOUR: () => {
                return this.#toDate(this.#gVal(stack[0])).getHours();
            },
            MINUTE: () => this.#toDate(this.#gVal(stack[0])).getMinutes(),

            SECOND: () => this.#toDate(this.#gVal(stack[0])).getSeconds(),

            WEEKNUM: () => {
                const date = this.#toDate(this.#gVal(stack[0]));
                const startOfYear = new Date(date.getFullYear(), 0, 1);
                const dayOfYear = ((date - startOfYear) / 86400000) + 1;
                return Math.ceil(dayOfYear / 7);
            },

            WEEKDAY: () => {
                const date = this.#toDate(this.#gVal(stack[0]));
                return date.getDay() + 1; // 返回1-7，其中1表示周日
            },

            EDATE: () => {
                this._lastResultIsDate = true;
                const startDate = this.#toDate(this.#gVal(stack[0]));
                const months = this.#gVal(stack[1]);
                startDate.setMonth(startDate.getMonth() + months);
                return dateToSerial(startDate);
            },

            EOMONTH: () => {
                this._lastResultIsDate = true;
                const startDate = this.#toDate(this.#gVal(stack[0]));
                const months = this.#gVal(stack[1]);
                startDate.setMonth(startDate.getMonth() + months + 1);
                startDate.setDate(0); // 将日期设为0表示上个月的最后一天
                return dateToSerial(startDate);
            },

            WORKDAY: () => {
                this._lastResultIsDate = true;
                const startDate = this.#toDate(this.#gVal(stack[0]));
                let days = this.#gVal(stack[1]);
                let currentDay = new Date(startDate);
                while (days > 0) {
                    currentDay.setDate(currentDay.getDate() + 1);
                    const dayOfWeek = currentDay.getDay();
                    if (dayOfWeek !== 0 && dayOfWeek !== 6) days--; // 跳过周末
                }
                return dateToSerial(currentDay);
            },

            WORKDAY_INTL: () => {
                this._lastResultIsDate = true;
                const startDate = this.#toDate(this.#gVal(stack[0]));
                let days = this.#gVal(stack[1]);
                const weekend = this.#gVal(stack[2]) || "0000011"; // 默认周末为周六、周日
                let currentDay = new Date(startDate);
                while (days > 0) {
                    currentDay.setDate(currentDay.getDate() + 1);
                    const dayOfWeek = currentDay.getDay();
                    if (weekend[dayOfWeek] === "0") days--; // 仅计数工作日
                }
                return dateToSerial(currentDay);
            },
            DAYS360: () => {
                const startDate = this.#toDate(this.#gVal(stack[0]));
                const endDate = this.#toDate(this.#gVal(stack[1]));
                const startDay = Math.min(startDate.getDate(), 30);
                const endDay = Math.min(endDate.getDate(), 30);
                return (endDate.getFullYear() - startDate.getFullYear()) * 360 +
                    (endDate.getMonth() - startDate.getMonth()) * 30 +
                    (endDay - startDay);
            },
            NETWORKDAYS: () => {
                const startDate = this.#toDate(this.#gVal(stack[0]));
                const endDate = this.#toDate(this.#gVal(stack[1]));
                const holidays = stack.slice(2).map(h => this.#toDate(this.#gVal(h))); // 假期列表
                let workDays = 0;

                for (let date = new Date(startDate); date <= endDate; date.setDate(date.getDate() + 1)) {
                    const day = date.getDay();
                    if (day !== 0 && day !== 6 && !holidays.some(h => h.getTime() === date.getTime())) {
                        workDays++;
                    }
                }
                return workDays;
            },

            NETWORKDAYS_INTL: () => {
                const startDate = this.#toDate(this.#gVal(stack[0]));
                const endDate = this.#toDate(this.#gVal(stack[1]));
                const weekendMask = this.#gVal(stack[2]) || '0000011'; // 默认周六日为周末
                const holidays = stack.slice(3).map(h => this.#toDate(this.#gVal(h))); // 假期列表
                let workDays = 0;

                for (let date = new Date(startDate); date <= endDate; date.setDate(date.getDate() + 1)) {
                    const day = date.getDay();
                    if (weekendMask[day] === '0' && !holidays.some(h => h.getTime() === date.getTime())) {
                        workDays++;
                    }
                }
                return workDays;
            },

            YEARFRAC: () => {
                const startDate = this.#toDate(this.#gVal(stack[0]));
                const endDate = this.#toDate(this.#gVal(stack[1]));
                const daysInYear = (startDate.getFullYear() % 4 === 0 && startDate.getFullYear() % 100 !== 0) || startDate.getFullYear() % 400 === 0 ? 366 : 365;
                const daysDiff = (endDate - startDate) / (1000 * 60 * 60 * 24);
                return daysDiff / daysInYear;
            },

            DATEDIF: () => {
                const startDate = this.#toDate(this.#gVal(stack[0]));
                const endDate = this.#toDate(this.#gVal(stack[1]));
                const unit = String(this.#gVal(stack[2])).toUpperCase();
                const diffInTime = endDate - startDate;

                switch (unit) {
                    case 'Y': return endDate.getFullYear() - startDate.getFullYear();
                    case 'M': return (endDate.getFullYear() - startDate.getFullYear()) * 12 + (endDate.getMonth() - startDate.getMonth());
                    case 'D': return Math.floor(diffInTime / (1000 * 60 * 60 * 24));
                    case 'YM': return (endDate.getMonth() - startDate.getMonth() + 12) % 12;
                    case 'YD': {
                        const tempDate = new Date(startDate);
                        tempDate.setFullYear(endDate.getFullYear());
                        if (tempDate > endDate) tempDate.setFullYear(endDate.getFullYear() - 1);
                        return Math.floor((endDate - tempDate) / (1000 * 60 * 60 * 24));
                    }
                    case 'MD': {
                        let days = endDate.getDate() - startDate.getDate();
                        if (days < 0) {
                            const prevMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 0);
                            days += prevMonth.getDate();
                        }
                        return days;
                    }
                    default: throw new Error("#VALUE! Invalid unit for DATEDIF.");
                }
            },

            CLEAN: () => String(this.#gVal(stack[0])).replace(/[\x00-\x1F\x7F]/g, ''),

            DOLLAR: () => {
                const number = this.#gVal(stack[0]);
                const decimals = this.#gVal(stack[1]) || 2;
                return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: decimals }).format(number);
            },

            FIXED: () => {
                const number = this.#gVal(stack[0]);
                const decimals = this.#gVal(stack[1]) || 2;
                const noCommas = this.#gVal(stack[2]) || false;
                const formattedNumber = number.toFixed(decimals);
                return noCommas ? formattedNumber.replace(/,/g, '') : formattedNumber;
            },

            TEXT: () => {
                let value = this.#gVal(stack[0]);
                const format = this.#gVal(stack[1]);

                // 使用 NumFmt 系统进行 Excel 风格的格式化
                // 支持日期格式如 "m.dd", "yyyy-mm-dd" 等
                if (value instanceof Date) {
                    value = dateTrans(value); // Date → Excel序列号
                }

                if (typeof value === 'number' || typeof value === 'string') {
                    return numFmtFun(value, format);
                }

                return String(value ?? '');
            },

            VALUE: () => {
                const val = this.#gVal(stack[0]);
                const parsed = parseFloat(val);
                return isNaN(parsed) ? null : parsed;
            },
            PROPER: () => {
                return String(this.#gVal(stack[0])).toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
            },
            CHAR: () => String.fromCharCode(this.#gVal(stack[0])),
            CODE: () => String(this.#gVal(stack[0])).charCodeAt(0),
            REPLACE: () => {
                const originalText = String(this.#gVal(stack[0]));
                const start = this.#gVal(stack[1]) - 1;
                const numChars = this.#gVal(stack[2]);
                const newText = String(this.#gVal(stack[3]));
                return originalText.slice(0, start) + newText + originalText.slice(start + numChars);
            },
            SUBSTITUTE: () => {
                const text = String(this.#gVal(stack[0]));
                const oldText = String(this.#gVal(stack[1]));
                const newText = String(this.#gVal(stack[2]));
                const occurrence = this.#gVal(stack[3]) ?? 0;
                if (occurrence === 0) {
                    return text.split(oldText).join(newText); // 替换所有
                }
                let count = 0;
                return text.replace(new RegExp(oldText, 'g'), (match) => {
                    count += 1;
                    return count === occurrence ? newText : match;
                });
            },
            REPT: () => String(this.#gVal(stack[0])).repeat(this.#gVal(stack[1])),
            FIND: () => {
                const findText = String(this.#gVal(stack[0]));
                const withinText = String(this.#gVal(stack[1]));
                const startPosition = this.#gVal(stack[2]) ? this.#gVal(stack[2]) - 1 : 0;
                const index = withinText.indexOf(findText, startPosition);
                return index >= 0 ? index + 1 : null;
            },
            SEARCH: () => {
                const findText = String(this.#gVal(stack[0]));
                const withinText = String(this.#gVal(stack[1]));
                const startNum = this.#gVal(stack[2]) ?? 1;
                return withinText.indexOf(findText, startNum - 1) + 1 || null;
            },
            EXACT: () => {
                return String(this.#gVal(stack[0])) === String(this.#gVal(stack[1]));
            },
            T: () => {
                const value = this.#gVal(stack[0]);
                return typeof value === 'string' ? value : '';
            },
            ISERROR: () => {
                return !!firstErrorValue(this.#gVal(stack[0]));
            },
            ISERR: () => {
                const err = firstErrorValue(this.#gVal(stack[0]));
                return !!err && errorCodeOf(err) !== '#N/A';
            },
            ISNA: () => {
                const err = firstErrorValue(this.#gVal(stack[0]));
                return !!err && errorCodeOf(err) === '#N/A';
            },
            'ERROR.TYPE': () => this.executeFunction('ERROR_TYPE', stack),
            ERROR_TYPE: () => {
                const type = errorTypeNumber(firstErrorValue(this.#gVal(stack[0])));
                if (type === null) throw new Error('#N/A');
                return type;
            },
            ISNUMBER: () => {
                return typeof this.#gVal(stack[0]) === 'number';
            },
            ISEVEN: () => {
                const value = this.#gVal(stack[0]);
                return typeof value === 'number' && value % 2 === 0;
            },
            ISODD: () => {
                const value = this.#gVal(stack[0]);
                return typeof value === 'number' && value % 2 !== 0;
            },
            N: () => {
                const value = this.#gVal(stack[0]);
                return typeof value === 'number' ? value : 0;
            },
            ISBLANK: () => {
                return this.#gVal(stack[0]) === null || this.#gVal(stack[0]) === '';
            },
            ISLOGICAL: () => typeof this.#gVal(stack[0]) === 'boolean',
            ISTEXT: () => typeof this.#gVal(stack[0]) === 'string',
            ISNONTEXT: () => typeof this.#gVal(stack[0]) !== 'string' && this.#gVal(stack[0]) != null,
            ISREF: () => Array.isArray(this.#gVal(stack[0])), // Assuming references are represented as arrays
            TYPE: () => {
                const value = this.#gVal(stack[0]);
                if (value === null || value === undefined) return 1; // Error
                if (typeof value === 'number') return 1; // Number
                if (typeof value === 'string') return 2; // Text
                if (typeof value === 'boolean') return 4; // Logical
                if (Array.isArray(value)) return 16; // Reference
                return 64; // Unknown type
            },
            NA: () => { throw new Error("#N/A"); },


            BIN2DEC: () => {
                const binary = this.#gVal(stack[0]);
                return parseInt(binary, 2);
            },
            BIN2HEX: () => parseInt(this.#gVal(stack[0]), 2).toString(16).toUpperCase(),
            BIN2OCT: () => parseInt(this.#gVal(stack[0]), 2).toString(8),
            DEC2BIN: () => parseInt(this.#gVal(stack[0]), 10).toString(2),
            DEC2HEX: () => parseInt(this.#gVal(stack[0]), 10).toString(16).toUpperCase(),
            DEC2OCT: () => parseInt(this.#gVal(stack[0]), 10).toString(8),
            HEX2BIN: () => parseInt(this.#gVal(stack[0]), 16).toString(2),
            HEX2DEC: () => parseInt(this.#gVal(stack[0]), 16).toString(10),
            HEX2OCT: () => parseInt(this.#gVal(stack[0]), 16).toString(8),
            OCT2BIN: () => parseInt(this.#gVal(stack[0]), 8).toString(2),
            OCT2DEC: () => parseInt(this.#gVal(stack[0]), 8).toString(10),
            OCT2HEX: () => parseInt(this.#gVal(stack[0]), 8).toString(16).toUpperCase(),
            ERF: () => {
                const x = this.#gVal(stack[0]);
                const erf = (z) => {
                    const t = 1 / (1 + 0.5 * Math.abs(z));
                    const tau = t * Math.exp(
                        -z * z - 1.26551223 +
                        t * (1.00002368 +
                            t * (0.37409196 +
                                t * (0.09678418 +
                                    t * (-0.18628806 +
                                        t * (0.27886807 +
                                            t * (-1.13520398 +
                                                t * (1.48851587 +
                                                    t * (-0.82215223 +
                                                        t * 0.17087277))))))))
                    );
                    return z >= 0 ? 1 - tau : tau - 1;
                };
                return erf(x);
            },
            ERFC: () => {
                const x = this.#gVal(stack[0]);
                const t = 1 / (1 + 0.5 * Math.abs(x));
                const tau = t * Math.exp(
                    -x * x - 1.26551223 +
                    t * (1.00002368 +
                        t * (0.37409196 +
                            t * (0.09678418 +
                                t * (-0.18628806 +
                                    t * (0.27886807 +
                                        t * (-1.13520398 +
                                            t * (1.48851587 +
                                                t * (-0.82215223 +
                                                    t * 0.17087277))))))))
                );
                return x >= 0 ? tau : 2 - tau;
            },
            ERFC_PRECISE: () => {
                const x = this.#gVal(stack[0]);
                const t = 1 / (1 + 0.5 * Math.abs(x));
                const tau = t * Math.exp(
                    -x * x - 1.26551223 +
                    t * (1.00002368 +
                        t * (0.37409196 +
                            t * (0.09678418 +
                                t * (-0.18628806 +
                                    t * (0.27886807 +
                                        t * (-1.13520398 +
                                            t * (1.48851587 +
                                                t * (-0.82215223 +
                                                    t * 0.17087277))))))))
                );
                return x >= 0 ? tau : 2 - tau;
            },
            DELTA: () => this.#gVal(stack[0]) === this.#gVal(stack[1]) ? 1 : 0,
            GESTEP: () => this.#gVal(stack[0]) >= (this.#gVal(stack[1]) || 0) ? 1 : 0,
            COMPLEX: () => {
                const real = this.#gVal(stack[0]);
                const imaginary = this.#gVal(stack[1]);
                const suffix = this.#gVal(stack[2]) || 'i';
                return `${real}${imaginary >= 0 ? '+' : ''}${imaginary}${suffix}`;
            },
            IMABS: () => {
                const complex = this.#gVal(stack[0]);
                const [real, imaginary] = complex.split(/(?=[+-])/).map(Number);
                return Math.sqrt(real * real + imaginary * imaginary);
            },
            IMAGINARY: () => {
                const complex = this.#gVal(stack[0]);
                return parseFloat(complex.split(/(?=[+-])/)[1]) || 0;
            },
            IMARGUMENT: () => {
                const complex = this.#gVal(stack[0]);
                const [real, imaginary] = complex.split(/(?=[+-])/).map(Number);
                return Math.atan2(imaginary, real);
            },
            IMCONJUGATE: () => {
                const complex = this.#gVal(stack[0]);
                const [real, imaginary] = complex.split(/(?=[+-])/).map(Number);
                return `${real}${imaginary >= 0 ? '-' : '+'}${Math.abs(imaginary)}i`;
            },
            IMCOS: () => {
                const complex = this.#gVal(stack[0]);
                const [real, imaginary] = complex.split(/(?=[+-])/).map(Number);
                return `${Math.cos(real) * Math.cosh(imaginary)}${-Math.sin(real) * Math.sinh(imaginary)}i`;
            },
            IMDIV: () => {
                const [num1, num2] = [this.#gVal(stack[0]), this.#gVal(stack[1])].map(c => c.split(/(?=[+-])/).map(Number));
                const [a, b, c, d] = [num1[0], num1[1], num2[0], num2[1]];
                const denominator = c * c + d * d;
                return `${(a * c + b * d) / denominator}${(b * c - a * d) / denominator >= 0 ? '+' : ''}${(b * c - a * d) / denominator}i`;
            },
            IMEXP: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                return math.exp(complex).toString();
            },
            IMLN: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                return math.log(complex).toString();
            },
            IMLOG10: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                return math.log10(complex).toString();
            },
            IMLOG2: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                return math.log2(complex).toString();
            },
            IMREAL: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                return complex.re;
            },
            IMSIN: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                return math.sin(complex).toString();
            },
            IMSQRT: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                return math.sqrt(complex).toString();
            },
            IMSUB: () => {
                const complex1 = math.complex(this.#gVal(stack[0]));
                const complex2 = math.complex(this.#gVal(stack[1]));
                return math.subtract(complex1, complex2).toString();
            },
            IMPOWER: () => {
                const complex = math.complex(this.#gVal(stack[0]));
                const power = this.#gVal(stack[1]);
                return math.pow(complex, power).toString();
            },
            IMPRODUCT: () => {
                const complexes = this.#gVals(stack).map(val => math.complex(val));
                return complexes.reduce((acc, complex) => math.multiply(acc, complex), math.complex(1)).toString();
            },
            IMSUM: () => {
                const complexes = this.#gVals(stack).map(val => math.complex(val));
                return complexes.reduce((acc, complex) => math.add(acc, complex), math.complex(0)).toString();
            },
            "RANK.AVG": () => {
                const number = this.#gVal(stack[0]);
                const array = this.#gNums(stack[1]);
                const rank = array.filter(x => x < number).length + 1;
                const duplicates = array.filter(x => x === number).length;
                return rank + (duplicates - 1) / 2;
            },
            COUNTIF: () => {
                const range = this.#gVals([stack[0]]);
                const criterion = this.#gVal(stack[1]);

                const test = (() => {
                    if (typeof criterion === 'number') {
                        return val => val === criterion;
                    }

                    if (typeof criterion === 'string') {
                        // 支持通配符
                        const wildcardPattern = criterion.replace(/([.+^=!:${}()|[\]\\])/g, '\\$1')
                            .replace(/\*/g, '.*')
                            .replace(/\?/g, '.');
                        if (/[*?]/.test(criterion)) {
                            const regex = new RegExp('^' + wildcardPattern + '$', 'i');
                            return val => {
                                if (val instanceof Error || typeof val === 'string' && val.startsWith('#')) {
                                    throw new Error('#REF!');
                                }
                                return typeof val === 'string' && regex.test(val);
                            };
                        }

                        // 处理逻辑符号
                        const m = criterion.match(/^([<>=!]=?|<>)(.*)$/);
                        if (m) {
                            let [, op, raw] = m;
                            if (op === '<>') op = '!=';

                            const num = parseFloat(raw);
                            const isNum = !isNaN(num);

                            return val => {
                                if (val instanceof Error || typeof val === 'string' && val.startsWith('#')) {
                                    throw new Error('#REF!');
                                }
                                const v = isNum ? parseFloat(val) : String(val);
                                switch (op) {
                                    case '>': return v > num;
                                    case '>=': return v >= num;
                                    case '<': return v < num;
                                    case '<=': return v <= num;
                                    case '=': return v == raw;
                                    case '!=': return v != raw;
                                    default: return false;
                                }
                            };
                        }

                        return val => {
                            if (val instanceof Error || typeof val === 'string' && val.startsWith('#')) {
                                throw new Error('#REF!');
                            }
                            return String(val).toLowerCase() === criterion.toLowerCase();
                        };
                    }

                    return val => {
                        if (val instanceof Error || typeof val === 'string' && val.startsWith('#')) {
                            throw new Error('#REF!');
                        }
                        return val == criterion;
                    };
                })();

                let count = 0;
                const rows = Array.isArray(range[0]) ? range : [range];
                for (const row of rows) {
                    for (const cell of row) {
                        if (cell instanceof Error || typeof cell === 'string' && cell.startsWith('#')) {
                            throw new Error('#REF!');
                        }
                        if (test(cell)) count++;
                    }
                }
                return count;
            }

        }
        const name = String(funcName || '').toUpperCase();
        const fn = F[name] ?? F[name.replace(/\./g, '_')];
        if (typeof fn !== 'function') throw new Error('#NAME?');
        return fn();
    }

}

function columnNameToIndex(name) {
    return String(name).toUpperCase().split('').reduce((sum, char) => {
        return sum * 26 + char.charCodeAt(0) - 64;
    }, 0) - 1;
}

