export const FORMULA_COMPATIBILITY_STATUS = {
    SUPPORTED: 'supported',
    PARTIAL: 'partial',
    UNSUPPORTED: 'unsupported'
};

const SUPPORTED = new Set([
    'ABS',
    'AND',
    'CHOOSE',
    'COUNT',
    'ERROR.TYPE',
    'FALSE',
    'IF',
    'IFERROR',
    'IFNA',
    'IFS',
    'INT',
    'ISERR',
    'ISERROR',
    'ISNA',
    'NA',
    'NOT',
    'OR',
    'ROUND',
    'SEQUENCE',
    'SUM',
    'SWITCH',
    'TEXT',
    'TRUE',
    'XOR'
]);

const PARTIAL = new Map([
    ['CELL', 'Core info types only; Excel format/protection fields are partial.'],
    ['DOLLAR', 'Uses browser Intl currency formatting; locale details may differ from Excel.'],
    ['FIND', 'Core behavior implemented; double-byte variants are partial.'],
    ['FINDB', 'Byte semantics are not fully locale-accurate.'],
    ['LEFTB', 'Delegates to LEFT; byte semantics are not fully locale-accurate.'],
    ['LENB', 'Delegates to LEN; byte semantics are not fully locale-accurate.'],
    ['MIDB', 'Delegates to MID; byte semantics are not fully locale-accurate.'],
    ['REPLACEB', 'Byte semantics are not fully locale-accurate.'],
    ['RIGHTB', 'Delegates to RIGHT; byte semantics are not fully locale-accurate.'],
    ['SEARCHB', 'Byte semantics are not fully locale-accurate.'],
    ['TEXTJOIN', 'Implemented, but not yet covered by Excel golden cases.'],
    ['OFFSET', 'Basic dynamic reference implemented; dependency graph treats it as volatile.'],
    ['INDIRECT', 'A1-style references implemented; dependency graph treats it as volatile.'],
    ['INFO', 'Returns SheetNext-compatible environment values, not native Excel host details.'],
    ['HYPERLINK', 'Display value is supported; link navigation side effects are outside formula evaluation.']
]);

const IMPLEMENTED_PARTIAL = new Set([
    'ACOS', 'ACOSH', 'ACOT', 'ACOTH', 'ADDRESS', 'AGGREGATE', 'ARABIC', 'AREAS',
    'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH', 'AVERAGE', 'AVERAGEA', 'AVERAGEIF',
    'AVERAGEIFS', 'AVEDEV', 'BASE', 'BIN2DEC', 'BIN2HEX', 'BIN2OCT', 'BITAND',
    'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'CEILING', 'CEILING.PRECISE',
    'CHAR', 'CHOOSECOLS', 'CHOOSEROWS', 'CLEAN', 'CODE', 'COLUMN', 'COLUMNS',
    'COMBIN', 'COMBINA', 'COMPLEX', 'CONCAT', 'CONCATENATE', 'CORREL', 'COS',
    'COSH', 'COT', 'COTH', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'COVAR', 'CSC',
    'CSCH', 'DATE', 'DATEDIF', 'DATEVALUE', 'DAY', 'DAYS', 'DAYS360', 'DB',
    'DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'DECIMAL', 'DEGREES', 'DELTA', 'DEVSQ',
    'DISC', 'DOLLARDE', 'DOLLARFR', 'DROP', 'EDATE', 'EFFECT', 'ENCODEURL',
    'EOMONTH', 'ERF', 'ERFC', 'ERFC.PRECISE', 'EVEN', 'EXACT', 'EXP', 'EXPAND',
    'FACT', 'FACTDOUBLE', 'FILTER', 'FIXED', 'FLOOR', 'FLOOR.PRECISE',
    'FORMULATEXT', 'FREQUENCY', 'FV', 'FVSCHEDULE', 'GAMMA', 'GAMMALN',
    'GAMMALN.PRECISE', 'GAUSS', 'GCD', 'GEOMEAN', 'GESTEP', 'GROWTH', 'HARMEAN',
    'GETPIVOTDATA', 'HEX2BIN', 'HEX2DEC', 'HEX2OCT', 'HLOOKUP', 'HOUR', 'HSTACK', 'IMABS',
    'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH', 'IMCOT',
    'IMCSC', 'IMCSCH', 'IMDIV', 'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMPOWER',
    'IMPRODUCT', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSQRT',
    'IMSUB', 'IMSUM', 'IMTAN', 'INDEX', 'INTERCEPT', 'IPMT', 'IRR', 'ISBLANK',
    'ISEVEN', 'ISFORMULA', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISODD',
    'ISOWEEKNUM', 'ISPMT', 'ISREF', 'ISTEXT', 'ISO.CEILING', 'KURT', 'LARGE',
    'LCM', 'LEFT', 'LEN', 'LINEST', 'LN', 'LOG', 'LOG10', 'LOGEST', 'LOOKUP',
    'LOWER', 'MATCH', 'MAX', 'MAXA', 'MAXIFS', 'MDETERM', 'MEDIAN', 'MID', 'MIN',
    'MINA', 'MINIFS', 'MINUTE', 'MINVERSE', 'MIRR', 'MMULT', 'MOD', 'MODE',
    'MROUND', 'MULTINOMIAL', 'MUNIT', 'NETWORKDAYS', 'NETWORKDAYS.INTL',
    'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'NORMDIST', 'NORMINV',
    'NORMSDIST', 'NORMSINV', 'NPER', 'NPV', 'NUMBERVALUE', 'OCT2BIN', 'OCT2DEC',
    'OCT2HEX', 'ODD', 'PERCENTILE', 'PERCENTILE.EXC', 'PERCENTILE.INC',
    'PERCENTRANK', 'PERCENTRANK.EXC', 'PERCENTRANK.INC', 'PERMUT', 'PERMUTATIONA',
    'PHI', 'PI', 'PMT', 'POWER', 'PPMT', 'PROB', 'PRODUCT', 'PROPER', 'PV',
    'QUARTILE', 'QUARTILE.EXC', 'QUARTILE.INC', 'QUOTIENT', 'RADIANS', 'RAND',
    'RANDARRAY', 'RANDBETWEEN', 'RANK', 'RANK.AVG', 'RANK.EQ', 'RATE', 'RECEIVED',
    'REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST', 'REPLACE', 'REPT', 'RIGHT',
    'ROMAN', 'ROUNDDOWN', 'ROUNDUP', 'ROW', 'ROWS', 'RSQ', 'SEARCH', 'SECOND',
    'SEC', 'SECH', 'SERIESSUM', 'SHEET', 'SHEETS', 'SIGN', 'SIN', 'SINH', 'SKEW',
    'SKEW.P', 'SLN', 'SLOPE', 'SMALL', 'SORT', 'SORTBY', 'SQRT', 'SQRTPI',
    'STANDARDIZE', 'STDEV', 'STDEV.P', 'STDEV.S', 'STDEVA', 'STDEVP', 'STDEVPA',
    'STEYX', 'SUBSTITUTE', 'SUBTOTAL', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ',
    'SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'SYD', 'T', 'TAKE', 'TAN', 'TANH',
    'TEXTAFTER', 'TEXTBEFORE', 'TEXTSPLIT', 'TIME', 'TIMEVALUE', 'TOCOL', 'TODAY',
    'TOROW', 'TRANSPOSE', 'TREND', 'TRIM', 'TRIMMEAN', 'TRIMRANGE', 'TRUNC',
    'TYPE', 'UNICHAR', 'UNICODE', 'UNIQUE', 'UPPER', 'VALUE', 'VALUETOTEXT',
    'VAR', 'VAR.P', 'VAR.S', 'VARA', 'VARP', 'VARPA', 'VDB', 'VLOOKUP', 'VSTACK',
    'WEEKDAY', 'WEEKNUM', 'WORKDAY', 'WORKDAY.INTL', 'WRAPCOLS', 'WRAPROWS',
    'XIRR', 'XLOOKUP', 'XMATCH', 'XNPV', 'YEAR', 'YEARFRAC', 'Z.TEST', 'ZTEST'
]);

const UNSUPPORTED = new Map([
    ['CALL', 'External procedures are intentionally unsupported.'],
    ['CUBE', 'Excel data model/cube functions are out of scope.'],
    ['CUBEKPIMEMBER', 'Excel data model/cube functions are out of scope.'],
    ['CUBEMEMBER', 'Excel data model/cube functions are out of scope.'],
    ['CUBEMEMBERPROPERTY', 'Excel data model/cube functions are out of scope.'],
    ['CUBERANKEDMEMBER', 'Excel data model/cube functions are out of scope.'],
    ['CUBESET', 'Excel data model/cube functions are out of scope.'],
    ['CUBESETCOUNT', 'Excel data model/cube functions are out of scope.'],
    ['CUBEVALUE', 'Excel data model/cube functions are out of scope.'],
    ['LAMBDA', 'Lambda evaluation engine is not implemented.'],
    ['LET', 'Scoped formula variable binding is not implemented.'],
    ['WEBSERVICE', 'Network formula calls are intentionally unsupported.']
]);

export function getFormulaCompatibility(name) {
    const key = normalizeFormulaName(name);
    if (hasName(SUPPORTED, key)) {
        return { status: FORMULA_COMPATIBILITY_STATUS.SUPPORTED, note: 'Covered by Excel golden cases.' };
    }

    const partialNote = getName(PARTIAL, key);
    if (partialNote) {
        return { status: FORMULA_COMPATIBILITY_STATUS.PARTIAL, note: partialNote };
    }

    if (hasName(IMPLEMENTED_PARTIAL, key)) {
        return {
            status: FORMULA_COMPATIBILITY_STATUS.PARTIAL,
            note: 'Implemented in the engine, but pending Excel golden-case verification.'
        };
    }

    const unsupportedNote = getName(UNSUPPORTED, key);
    if (unsupportedNote) {
        return { status: FORMULA_COMPATIBILITY_STATUS.UNSUPPORTED, note: unsupportedNote };
    }

    return {
        status: FORMULA_COMPATIBILITY_STATUS.UNSUPPORTED,
        note: 'No compatibility claim yet; add Excel golden cases before enabling.'
    };
}

export function summarizeFormulaCompatibility(formulas) {
    return formulas.reduce((stats, item) => {
        const { status } = getFormulaCompatibility(item.name);
        stats[status] = (stats[status] || 0) + 1;
        stats.total++;
        return stats;
    }, {
        total: 0,
        supported: 0,
        partial: 0,
        unsupported: 0
    });
}

function normalizeFormulaName(name) {
    return String(name || '').trim().toUpperCase();
}

function aliasNames(name) {
    const aliases = [name];
    if (name.includes('_')) aliases.push(name.replace(/_/g, '.'));
    if (name.includes('.')) aliases.push(name.replace(/\./g, '_'));
    return aliases;
}

function hasName(set, name) {
    return aliasNames(name).some(alias => set.has(alias));
}

function getName(map, name) {
    for (const alias of aliasNames(name)) {
        if (map.has(alias)) return map.get(alias);
    }
    return null;
}

export default {
    FORMULA_COMPATIBILITY_STATUS,
    getFormulaCompatibility,
    summarizeFormulaCompatibility
};
