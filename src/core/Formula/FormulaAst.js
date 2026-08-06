export const FORMULA_AST_VOLATILE_FUNCTIONS = new Set([
    'RAND',
    'RANDARRAY',
    'RANDBETWEEN',
    'NOW',
    'TODAY',
    'OFFSET',
    'INDIRECT',
    'INFO',
    'CELL',
    'SUBTOTAL'
]);

const BINARY_PRECEDENCE = {
    ':': 8,
    '^': 6,
    '*': 5,
    '/': 5,
    '+': 4,
    '-': 4,
    '&': 3,
    '=': 2,
    '<': 2,
    '>': 2,
    '<=': 2,
    '>=': 2,
    '<>': 2
};

const RIGHT_ASSOCIATIVE = new Set(['^']);
const UNARY_PRECEDENCE = 7;

export class FormulaParseError extends Error {
    constructor(message) {
        super(message);
        this.name = 'FormulaParseError';
    }
}

export function parseFormula(formula) {
    return new FormulaAstParser(formula).parse();
}

/**
 * Tolerantly tokenize a formula input for editing support.
 * Accepts the full input including the leading '=', never throws on
 * incomplete/invalid input, and returns tokens carrying their
 * [start, end) offsets within the original input string.
 * @param {string} text - Full formula input, e.g. "=SUM(A1,"
 * @returns {Array<{type:string,value?:any,start:number,end:number}>}
 */
export function tokenizeFormula(text) {
    const raw = String(text ?? '');
    const offset = raw.startsWith('=') ? 1 : 0;
    const tokens = new FormulaLexer(raw.slice(offset), { tolerant: true }).scan();
    if (offset) {
        tokens.forEach(token => {
            token.start += offset;
            token.end += offset;
        });
    }
    return tokens;
}

export function hasVolatileFunctionInAst(ast) {
    let found = false;
    walkAst(ast, node => {
        if (node.type === 'CallExpression' && FORMULA_AST_VOLATILE_FUNCTIONS.has(node.name)) {
            found = true;
            return false;
        }
        return true;
    });
    return found;
}

export function collectReferenceTexts(ast) {
    const refs = [];
    walkAst(ast, node => {
        if (node.type === 'Reference') refs.push(node.ref);
        else if (node.type === 'RangeExpression') {
            const text = rangeExpressionToText(node);
            if (text) refs.push(text);
            return false;
        }
        return true;
    });
    return refs;
}

export function isReferenceText(text) {
    const { address } = splitSheetReference(text);
    return isAddressText(address);
}

export function splitSheetReference(text) {
    const raw = String(text ?? '').trim();
    const bang = findSheetSeparator(raw);
    if (bang < 0) return { sheetName: null, address: raw };

    let sheetName = raw.slice(0, bang);
    const address = raw.slice(bang + 1);
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
        sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
    }
    return { sheetName, address };
}

export function walkAst(node, visit) {
    if (!node || visit(node) === false) return;

    switch (node.type) {
        case 'UnaryExpression':
        case 'PercentExpression':
        case 'ImplicitIntersectionExpression':
        case 'SpillReference':
            walkAst(node.argument, visit);
            break;
        case 'BinaryExpression':
        case 'RangeExpression':
            walkAst(node.left, visit);
            walkAst(node.right, visit);
            break;
        case 'CallExpression':
            node.args.forEach(arg => walkAst(arg, visit));
            break;
        case 'ArrayExpression':
            node.rows.forEach(row => row.forEach(item => walkAst(item, visit)));
            break;
    }
}

class FormulaAstParser {
    constructor(formula) {
        const text = String(formula ?? '');
        this._formula = text.startsWith('=') ? text.slice(1) : text;
        this._tokens = new FormulaLexer(this._formula).scan();
        this._pos = 0;
    }

    parse() {
        const expression = this._parseExpression(0);
        this._expect('EOF');
        return expression;
    }

    _parseExpression(minPrecedence) {
        let left = this._parsePrefix();

        while (true) {
            const token = this._peek();
            if (token.type === 'OPERATOR' && token.value === '%') {
                this._next();
                left = { type: 'PercentExpression', argument: left };
                continue;
            }
            if (token.type === 'HASH') {
                this._next();
                left = { type: 'SpillReference', argument: left };
                continue;
            }

            if (token.type !== 'OPERATOR') break;
            const precedence = BINARY_PRECEDENCE[token.value];
            if (precedence === undefined || precedence < minPrecedence) break;

            this._next();
            const nextMin = RIGHT_ASSOCIATIVE.has(token.value) ? precedence : precedence + 1;
            const right = this._parseExpression(nextMin);
            left = token.value === ':'
                ? { type: 'RangeExpression', left, right }
                : { type: 'BinaryExpression', operator: token.value, left, right };
        }

        return left;
    }

    _parsePrefix() {
        const token = this._next();

        if (token.type === 'OPERATOR' && (token.value === '+' || token.value === '-')) {
            return {
                type: 'UnaryExpression',
                operator: token.value,
                argument: this._parseExpression(UNARY_PRECEDENCE)
            };
        }
        if (token.type === 'AT') {
            return {
                type: 'ImplicitIntersectionExpression',
                argument: this._parseExpression(UNARY_PRECEDENCE)
            };
        }

        switch (token.type) {
            case 'NUMBER':
                return { type: 'Literal', value: token.value };
            case 'STRING':
                return { type: 'Literal', value: token.value };
            case 'IDENT':
                return this._parseIdentifier(token.value);
            case 'LPAREN': {
                const expression = this._parseExpression(0);
                this._expect('RPAREN');
                return expression;
            }
            case 'LBRACE':
                return this._parseArray();
            case 'COMMA':
            case 'SEMICOLON':
            case 'RPAREN':
            case 'RBRACE':
                this._pos--;
                return { type: 'Blank' };
            default:
                throw new FormulaParseError(`Unexpected token ${token.type}`);
        }
    }

    _parseIdentifier(identifier) {
        const upper = identifier.toUpperCase();
        if (this._peek().type === 'LPAREN') {
            this._next();
            return this._parseCall(upper);
        }
        if (upper === 'TRUE') return { type: 'Literal', value: true };
        if (upper === 'FALSE') return { type: 'Literal', value: false };
        if (isReferenceText(identifier)) return { type: 'Reference', ref: identifier };
        return { type: 'Name', name: identifier };
    }

    _parseCall(name) {
        const args = [];
        if (this._peek().type === 'RPAREN') {
            this._next();
            return { type: 'CallExpression', name, args };
        }

        while (true) {
            if (this._peek().type === 'COMMA') {
                args.push({ type: 'Blank' });
                this._next();
            } else {
                args.push(this._parseExpression(0));
            }

            if (this._peek().type === 'COMMA') {
                this._next();
                if (this._peek().type === 'RPAREN') args.push({ type: 'Blank' });
                continue;
            }
            this._expect('RPAREN');
            return { type: 'CallExpression', name, args };
        }
    }

    _parseArray() {
        const rows = [[]];
        if (this._peek().type === 'RBRACE') {
            this._next();
            return { type: 'ArrayExpression', rows };
        }

        while (true) {
            rows[rows.length - 1].push(this._parseExpression(0));

            const token = this._peek();
            if (token.type === 'COMMA') {
                this._next();
                continue;
            }
            if (token.type === 'SEMICOLON') {
                this._next();
                rows.push([]);
                continue;
            }
            this._expect('RBRACE');
            return { type: 'ArrayExpression', rows };
        }
    }

    _peek() {
        return this._tokens[this._pos] ?? { type: 'EOF' };
    }

    _next() {
        return this._tokens[this._pos++] ?? { type: 'EOF' };
    }

    _expect(type) {
        const token = this._next();
        if (token.type !== type) throw new FormulaParseError(`Expected ${type}, got ${token.type}`);
        return token;
    }
}

class FormulaLexer {
    constructor(formula, options = {}) {
        this._formula = formula;
        this._pos = 0;
        this._tokens = [];
        this._tolerant = options.tolerant === true;
        this._tokenStart = 0;
        this._unterminated = false;
    }

    scan() {
        while (this._pos < this._formula.length) {
            const ch = this._formula[this._pos];

            if (/\s/.test(ch)) {
                this._pos++;
                continue;
            }

            this._tokenStart = this._pos;

            if (ch === '"') {
                this._push('STRING', this._readQuoted('"'));
                continue;
            }

            if (ch === "'") {
                this._readSingleQuoted();
                continue;
            }

            if (isDigit(ch) || (ch === '.' && isDigit(this._formula[this._pos + 1]))) {
                this._push('NUMBER', this._readNumber());
                continue;
            }

            if (isIdentifierStart(ch)) {
                this._push('IDENT', this._readIdentifier());
                continue;
            }

            if (ch === '(') {
                this._pos++;
                this._push('LPAREN');
                continue;
            }
            if (ch === ')') {
                this._pos++;
                this._push('RPAREN');
                continue;
            }
            if (ch === '{') {
                this._pos++;
                this._push('LBRACE');
                continue;
            }
            if (ch === '}') {
                this._pos++;
                this._push('RBRACE');
                continue;
            }
            if (ch === ',') {
                this._pos++;
                this._push('COMMA');
                continue;
            }
            if (ch === ';') {
                this._pos++;
                this._push('SEMICOLON');
                continue;
            }
            if (ch === '@') {
                this._pos++;
                this._push('AT');
                continue;
            }
            if (ch === '#') {
                this._pos++;
                this._push('HASH');
                continue;
            }

            const op = this._readOperator();
            if (op) {
                this._push('OPERATOR', op);
                continue;
            }

            if (this._tolerant) {
                this._pos++;
                this._push('UNKNOWN', ch);
                continue;
            }
            throw new FormulaParseError(`Unexpected character ${ch}`);
        }

        this._tokenStart = this._pos;
        this._push('EOF');
        return this._tokens;
    }

    _push(type, value) {
        const token = { type, start: this._tokenStart, end: this._pos };
        if (value !== undefined) token.value = value;
        this._tokens.push(token);
    }

    _readQuoted(quote) {
        this._pos++;
        let value = '';
        while (this._pos < this._formula.length) {
            const ch = this._formula[this._pos];
            if (ch === quote) {
                if (this._formula[this._pos + 1] === quote) {
                    value += quote;
                    this._pos += 2;
                    continue;
                }
                this._pos++;
                return value;
            }
            value += ch;
            this._pos++;
        }
        if (this._tolerant) {
            this._unterminated = true;
            return value;
        }
        throw new FormulaParseError('Unterminated string literal');
    }

    _readSingleQuoted() {
        this._unterminated = false;
        const start = this._pos;
        const quoted = this._readQuoted("'");
        if (!this._unterminated && this._formula[this._pos] === '!') {
            this._pos++;
            const address = this._readIdentifierBody();
            this._push('IDENT', `'${quoted.replace(/'/g, "''")}'!${address}`);
            return;
        }
        this._push('STRING', this._unterminated ? quoted : this._formula.slice(start + 1, this._pos - 1));
    }

    _readNumber() {
        const start = this._pos;
        while (isDigit(this._formula[this._pos])) this._pos++;
        if (this._formula[this._pos] === '.') {
            this._pos++;
            while (isDigit(this._formula[this._pos])) this._pos++;
        }
        const e = this._formula[this._pos];
        if (e === 'e' || e === 'E') {
            const save = this._pos;
            this._pos++;
            if (this._formula[this._pos] === '+' || this._formula[this._pos] === '-') this._pos++;
            const expStart = this._pos;
            while (isDigit(this._formula[this._pos])) this._pos++;
            if (expStart === this._pos) this._pos = save;
        }
        return Number(this._formula.slice(start, this._pos));
    }

    _readIdentifier() {
        return this._readIdentifierBody();
    }

    _readIdentifierBody() {
        const start = this._pos;
        while (this._pos < this._formula.length && isIdentifierPart(this._formula[this._pos])) {
            this._pos++;
        }
        return this._formula.slice(start, this._pos);
    }

    _readOperator() {
        const two = this._formula.slice(this._pos, this._pos + 2);
        if (two === '<=' || two === '>=' || two === '<>' || two === '==') {
            this._pos += 2;
            return two === '==' ? '=' : two;
        }
        const one = this._formula[this._pos];
        if ('+-*/^&=<>%:'.includes(one)) {
            this._pos++;
            return one;
        }
        return null;
    }
}

function rangeExpressionToText(node) {
    const left = referenceTextFromNode(node.left);
    const right = referenceTextFromNode(node.right);
    if (!left || !right) return null;

    const leftRef = splitSheetReference(left);
    const rightRef = splitSheetReference(right);
    const sheet = rightRef.sheetName ?? leftRef.sheetName;
    return `${sheet ? quoteSheetName(sheet) + '!' : ''}${leftRef.address}:${rightRef.address}`;
}

function referenceTextFromNode(node) {
    if (node.type === 'Reference') return node.ref;
    if (node.type === 'RangeExpression') return rangeExpressionToText(node);
    return null;
}

function quoteSheetName(sheetName) {
    return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)
        ? sheetName
        : `'${String(sheetName).replace(/'/g, "''")}'`;
}

function findSheetSeparator(text) {
    let inQuote = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === "'") {
            if (inQuote && text[i + 1] === "'") {
                i++;
            } else {
                inQuote = !inQuote;
            }
        } else if (ch === '!' && !inQuote) {
            return i;
        }
    }
    return -1;
}

function isAddressText(address) {
    const cell = '\\$?[A-Z]{1,3}\\$?\\d+';
    const col = '\\$?[A-Z]{1,3}';
    const row = '\\$?\\d+';
    return new RegExp(`^${cell}$`, 'i').test(address)
        || new RegExp(`^${cell}:${cell}$`, 'i').test(address)
        || new RegExp(`^${col}:${col}$`, 'i').test(address)
        || new RegExp(`^${row}:${row}$`, 'i').test(address);
}

function isIdentifierStart(ch) {
    return /[A-Za-z_$]/.test(ch) || (ch >= '\u4e00' && ch <= '\u9fa5');
}

function isIdentifierPart(ch) {
    return /[A-Za-z0-9_$!.:]/.test(ch) || (ch >= '\u4e00' && ch <= '\u9fa5');
}

function isDigit(ch) {
    return ch >= '0' && ch <= '9';
}
