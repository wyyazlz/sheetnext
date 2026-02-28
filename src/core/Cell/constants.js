/**
 * 项目常量定义
 * 存放Excel相关的枚举和常量配置
 */

/**
 * Excel 单元格类型映射
 */
export const CELL_TYPES = {
    'n': 'number',
    's': 'string',
    'inlineStr': 'string',
    'str': 'string',
    'b': 'boolean',
    'e': 'error',
    'd': 'date'
};

/**
 * Excel 内置数字格式
 * @see https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat
 */
export const BUILTIN_NUMFMTS = {
    "0": 'General',
    "1": '0',
    "2": '0.00',
    "3": '#,##0',
    "4": '#,##0.00',
    "9": '0%',
    "10": '0.00%',
    "11": '0.00E+00',
    "12": '# ?/?',
    "13": '# ??/??',
    "14": 'yyyy/m/d',
    "15": 'd-mmm-yy',
    "16": 'd-mmm',
    "17": 'mmm-yy',
    "18": 'h:mm AM/PM',
    "19": 'h:mm:ss AM/PM',
    "20": 'h:mm',
    "21": 'h:mm:ss',
    "22": 'yyyy/m/d h:mm',
    "37": '#,##0 ;(#,##0)',
    "38": '#,##0 ;[Red](#,##0)',
    "39": '#,##0.00;(#,##0.00)',
    "40": '#,##0.00;[Red](#,##0.00)',
    "45": 'mm:ss',
    "46": '[h]:mm:ss',
    "47": 'mmss.0',
    "48": '##0.0E+0',
    "49": '@',
    // 中文格式
    '27': 'yyyy"年"m"月"',
    '28': 'm"月"d"日"',
    '29': 'm"月"d"日"',
    '30': 'm-d-yy',
    '31': 'yyyy"年"m"月"d"日"',
    '32': 'h"时"mm"分"',
    '33': 'h"时"mm"分"ss"秒"',
    '34': '上午/下午h"时"mm"分"',
    '35': '上午/下午h"时"mm"分"ss"秒"',
    '36': 'yyyy"年"m"月"',
    '50': 'yyyy"年"m"月"',
    '51': 'm"月"d"日"',
    '52': 'yyyy"年"m"月"',
    '53': 'm"月"d"日"',
    '54': 'm"月"d"日"',
    '55': '上午/下午h"时"mm"分"',
    '56': '上午/下午h"时"mm"分"ss"秒"',
    '57': 'yyyy"年"m"月"',
    '58': 'm"月"d"日"'
};
