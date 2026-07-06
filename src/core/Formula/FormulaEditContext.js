/**
 * Formula Edit Context Module
 * Pure text analysis over the tolerant formula token stream.
 * Powers the inline formula editing experience: function autocomplete,
 * parameter hints, colored reference highlighting and reference pointing.
 */
import { tokenizeFormula, isReferenceText } from './FormulaAst.js';

const FUNCTION_PREFIX_RE = /^[A-Za-z][A-Za-z0-9_.]*$/;
const CELL_PART_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;
const COL_PART_RE = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_PART_RE = /^(\$?)(\d+)$/;

/**
 * Analyze a formula input string for editing support.
 * @param {string} text - Full editor text, must start with '='
 * @param {number|null} [cursor] - Caret offset within text, null when unknown
 * @returns {Object|null} Analysis result, or null when text is not a formula
 *   - tokens: tolerant token stream with offsets
 *   - refs: reference tokens `{text, start, end, colorIndex}`, same text shares a color index
 *   - fnPrefix: identifier being typed at cursor `{text, start, end}` for autocomplete
 *   - call: innermost function call at cursor `{name, argIndex}` for the parameter hint
 *   - canInsertRef: whether a reference can be inserted/pointed at the cursor
 *   - refAtCursor: complete reference token the cursor sits in, pointing replaces it
 */
export function analyzeFormulaInput(text, cursor = null) {
    const raw = String(text ?? '');
    if (!raw.startsWith('=')) return null;

    const tokens = tokenizeFormula(raw);
    const refs = [];
    const colorByKey = new Map();
    tokens.forEach((token, i) => {
        if (token.type !== 'IDENT' || tokens[i + 1]?.type === 'LPAREN') return;
        if (!isReferenceText(token.value)) return;
        const key = String(token.value).toUpperCase();
        if (!colorByKey.has(key)) colorByKey.set(key, colorByKey.size);
        refs.push({
            text: String(token.value),
            start: token.start,
            end: token.end,
            colorIndex: colorByKey.get(key)
        });
    });

    const result = {
        isFormula: true,
        tokens,
        refs,
        fnPrefix: null,
        call: null,
        canInsertRef: false,
        refAtCursor: null
    };
    if (cursor == null || cursor < 1) return result;

    // Token the cursor sits inside. The end boundary only counts as inside
    // for IDENT tokens (typing a name/reference), delimiters and literals
    // ending at the cursor are treated as "before" it instead.
    let inside = null;
    let insideIndex = -1;
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (token.type === 'EOF') break;
        if (token.start >= cursor) break;
        if (cursor < token.end || (cursor === token.end && token.type === 'IDENT')) {
            inside = token;
            insideIndex = i;
            break;
        }
    }

    if (inside?.type === 'IDENT' && tokens[insideIndex + 1]?.type !== 'LPAREN') {
        const prefix = raw.slice(inside.start, cursor);
        if (FUNCTION_PREFIX_RE.test(prefix)) {
            result.fnPrefix = { text: prefix, start: inside.start, end: inside.end };
        }
        if (isReferenceText(inside.value)) {
            result.refAtCursor = { text: String(inside.value), start: inside.start, end: inside.end };
        }
    }

    // Innermost unclosed function call before the cursor
    const stack = [];
    let prev = null;
    for (const token of tokens) {
        if (token.type === 'EOF' || token.start >= cursor) break;
        switch (token.type) {
            case 'LPAREN':
                stack.push({ name: prev?.type === 'IDENT' ? String(prev.value).toUpperCase() : null, argIndex: 0 });
                break;
            case 'RPAREN':
                stack.pop();
                break;
            case 'LBRACE':
                stack.push({ name: null, argIndex: 0 });
                break;
            case 'RBRACE':
                stack.pop();
                break;
            case 'COMMA':
            case 'SEMICOLON':
                if (stack.length) stack[stack.length - 1].argIndex++;
                break;
        }
        prev = token;
    }
    for (let i = stack.length - 1; i >= 0; i--) {
        if (stack[i].name) {
            result.call = { name: stack[i].name, argIndex: stack[i].argIndex };
            break;
        }
    }

    // Reference insert position: after '=', an operator, '(', ',' or ';',
    // or replacing the complete reference the cursor sits in
    if (result.refAtCursor) {
        result.canInsertRef = true;
    } else if (!inside) {
        let before = null;
        for (const token of tokens) {
            if (token.type === 'EOF' || token.end > cursor) break;
            before = token;
        }
        if (!before) {
            result.canInsertRef = true;
        } else if (before.type === 'OPERATOR' && before.value !== '%') {
            result.canInsertRef = true;
        } else if (before.type === 'LPAREN' || before.type === 'COMMA' || before.type === 'SEMICOLON') {
            result.canInsertRef = true;
        }
    }

    return result;
}

/**
 * Cycle the $ anchors of a reference like Excel's F4 key.
 * A1 -> $A$1 -> A$1 -> $A1 -> A1; column/row-only parts toggle their anchor.
 * @param {string} refText - Reference text, may carry a sheet prefix
 * @returns {string} Reference with the next anchor state, or the input when not cyclable
 */
export function cycleReferenceAnchor(refText) {
    const raw = String(refText ?? '');
    const bang = raw.lastIndexOf('!');
    const prefix = bang >= 0 ? raw.slice(0, bang + 1) : '';
    const address = bang >= 0 ? raw.slice(bang + 1) : raw;
    const parts = address.split(':');
    if (!parts.length || parts.length > 2) return raw;

    const state = _partAnchorState(parts[0]);
    if (state === null) return raw;
    const next = (state + 1) % 4;
    const out = parts.map(part => _applyPartAnchor(part, next));
    if (out.some(part => part === null)) return raw;
    return prefix + out.join(':');
}

// Anchor states: 0=A1, 1=$A$1, 2=A$1, 3=$A1
function _partAnchorState(part) {
    let match = part.match(CELL_PART_RE);
    if (match) {
        const colAbs = match[1] === '$';
        const rowAbs = match[3] === '$';
        if (colAbs && rowAbs) return 1;
        if (rowAbs) return 2;
        if (colAbs) return 3;
        return 0;
    }
    match = part.match(COL_PART_RE);
    if (match) return match[1] === '$' ? 1 : 0;
    match = part.match(ROW_PART_RE);
    if (match) return match[1] === '$' ? 2 : 0;
    return null;
}

function _applyPartAnchor(part, state) {
    const colAbs = state === 1 || state === 3;
    const rowAbs = state === 1 || state === 2;
    let match = part.match(CELL_PART_RE);
    if (match) return `${colAbs ? '$' : ''}${match[2]}${rowAbs ? '$' : ''}${match[4]}`;
    match = part.match(COL_PART_RE);
    if (match) return `${colAbs ? '$' : ''}${match[2]}`;
    match = part.match(ROW_PART_RE);
    if (match) return `${rowAbs ? '$' : ''}${match[2]}`;
    return null;
}
