/**
 * Rule Factory Function
 * Separate files to avoid circular dependencies
 */
import { CFRule } from './CFRule.js';
import { CellIsRule } from './rules/CellIsRule.js';
import { ColorScaleRule } from './rules/ColorScaleRule.js';
import { DataBarRule } from './rules/DataBarRule.js';
import { IconSetRule } from './rules/IconSetRule.js';
import { TextRule } from './rules/TextRule.js';
import { DuplicateRule } from './rules/DuplicateRule.js';
import { Top10Rule } from './rules/Top10Rule.js';
import { AboveAverageRule } from './rules/AboveAverageRule.js';
import { TimePeriodRule } from './rules/TimePeriodRule.js';
import { BlankErrorRule } from './rules/BlankErrorRule.js';
import { ExpressionRule } from './rules/ExpressionRule.js';

/**
 * Rule Factory Function
 * @param {Object} config - Rule Configuration
 * @param {Object} sheet - Sheet instance
 * @ returns {CFRule} rule instance
 */
export function createRule(config, sheet) {
    const type = config.type;

    switch (type) {
        case 'cellIs':
            return new CellIsRule(config, sheet);

        case 'colorScale':
            return new ColorScaleRule(config, sheet);

        case 'dataBar':
            return new DataBarRule(config, sheet);

        case 'iconSet':
            return new IconSetRule(config, sheet);

        case 'containsText':
        case 'notContainsText':
        case 'beginsWith':
        case 'endsWith':
            return new TextRule(config, sheet);

        case 'duplicateValues':
        case 'uniqueValues':
            return new DuplicateRule(config, sheet);

        case 'top10':
            return new Top10Rule(config, sheet);

        case 'aboveAverage':
            return new AboveAverageRule(config, sheet);

        case 'timePeriod':
            return new TimePeriodRule(config, sheet);

        case 'containsBlanks':
        case 'notContainsBlanks':
        case 'containsErrors':
        case 'notContainsErrors':
            return new BlankErrorRule(config, sheet);

        case 'expression':
            return new ExpressionRule(config, sheet);

        default:
            // 未知类型，使用基类
            return new CFRule(config, sheet);
    }
}
