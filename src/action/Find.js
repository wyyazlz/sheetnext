/**
 * 查找相关操作
 */

import FindReplaceDialog from '../components/FindReplaceDialog.js';

let findDialogInstance = null;

/**
 * 打开查找替换定位对话框
 * @param {string} tab - 初始标签: find | replace | goto
 */
export function openFindDialog(tab = 'find') {
    if (!findDialogInstance) {
        findDialogInstance = new FindReplaceDialog(this.SN);
    }
    findDialogInstance.show(tab);
}
