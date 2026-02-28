/**
 * Text规则 - 文本相关
 * 包含：containsText, notContainsText, beginsWith, endsWith
 */
import { CFRule } from '../CFRule.js';

export class TextRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
        this.text = config.text || '';
    }

    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (val === null || val === undefined) return false;

        const cellText = String(val);
        const searchText = String(this.text);

        switch (this.type) {
            case 'containsText':
                return cellText.includes(searchText);

            case 'notContainsText':
                return !cellText.includes(searchText);

            case 'beginsWith':
                return cellText.startsWith(searchText);

            case 'endsWith':
                return cellText.endsWith(searchText);

            default:
                return false;
        }
    }

    _addXmlAttributes(node, cf) {
        const text = String(this.text ?? '');
        node['_$text'] = text;

        const firstRange = String(this.sqref || '').split(/\s+/).find(Boolean) || 'A1';
        let baseRef = 'A1';
        try {
            if (this.sheet?.rangeStrToNum && this.sheet?.Utils?.numToChar) {
                const range = this.sheet.rangeStrToNum(firstRange);
                baseRef = `${this.sheet.Utils.numToChar(range.s.c)}${range.s.r + 1}`;
            }
        } catch (e) {
            baseRef = 'A1';
        }

        const escaped = text.replace(/"/g, '""');
        switch (this.type) {
            case 'containsText':
                node['_$operator'] = 'containsText';
                node.formula = `NOT(ISERROR(SEARCH("${escaped}",${baseRef})))`;
                break;
            case 'notContainsText':
                node['_$operator'] = 'notContains';
                node.formula = `ISERROR(SEARCH("${escaped}",${baseRef}))`;
                break;
            case 'beginsWith':
                node['_$operator'] = 'beginsWith';
                node.formula = `LEFT(${baseRef},LEN("${escaped}"))="${escaped}"`;
                break;
            case 'endsWith':
                node['_$operator'] = 'endsWith';
                node.formula = `RIGHT(${baseRef},LEN("${escaped}"))="${escaped}"`;
                break;
            default:
                break;
        }
    }
}
