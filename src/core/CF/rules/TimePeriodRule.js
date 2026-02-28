/**
 * TimePeriod规则 - 日期时间段
 * 支持：today, yesterday, tomorrow, last7Days, lastWeek, thisWeek, nextWeek, lastMonth, thisMonth, nextMonth
 */
import { CFRule } from '../CFRule.js';
import { getTodayStart, getWeekStart, getMonthStart, excelDateToTimestamp } from '../helpers.js';

const DAY_MS = 86400000;
const WEEK_MS = 7 * DAY_MS;

export class TimePeriodRule extends CFRule {
    constructor(config, sheet) {
        super(config, sheet);
        this.timePeriod = config.timePeriod || 'today';
    }

    evaluate(cell, r, c, rangeData) {
        const val = cell.calcVal;
        if (typeof val !== 'number' || isNaN(val)) return false;

        // 将Excel日期序列号转为时间戳
        const cellTimestamp = excelDateToTimestamp(val);
        const cellDate = new Date(cellTimestamp);
        const cellDayStart = new Date(cellDate.getFullYear(), cellDate.getMonth(), cellDate.getDate()).getTime();

        const today = getTodayStart();
        const weekStart = getWeekStart();
        const monthStart = getMonthStart();
        const now = new Date();

        switch (this.timePeriod) {
            case 'today':
                return cellDayStart === today;

            case 'yesterday':
                return cellDayStart === today - DAY_MS;

            case 'tomorrow':
                return cellDayStart === today + DAY_MS;

            case 'last7Days':
                return cellDayStart >= today - 6 * DAY_MS && cellDayStart <= today;

            case 'lastWeek':
                return cellDayStart >= weekStart - WEEK_MS && cellDayStart < weekStart;

            case 'thisWeek':
                return cellDayStart >= weekStart && cellDayStart < weekStart + WEEK_MS;

            case 'nextWeek':
                return cellDayStart >= weekStart + WEEK_MS && cellDayStart < weekStart + 2 * WEEK_MS;

            case 'lastMonth': {
                const lastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                const lastMonthStart = lastMonth.getTime();
                const lastMonthEnd = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
                return cellDayStart >= lastMonthStart && cellDayStart < lastMonthEnd;
            }

            case 'thisMonth': {
                const nextMonthStart = new Date(now.getFullYear(), now.getMonth() + 1, 1).getTime();
                return cellDayStart >= monthStart && cellDayStart < nextMonthStart;
            }

            case 'nextMonth': {
                const nextMonthStart = new Date(now.getFullYear(), now.getMonth() + 1, 1).getTime();
                const nextNextMonthStart = new Date(now.getFullYear(), now.getMonth() + 2, 1).getTime();
                return cellDayStart >= nextMonthStart && cellDayStart < nextNextMonthStart;
            }

            default:
                return false;
        }
    }

    _addXmlAttributes(node, cf) {
        node['_$timePeriod'] = this.timePeriod;
    }
}
