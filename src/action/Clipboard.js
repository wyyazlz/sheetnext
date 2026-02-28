/**
 * Action 层剪贴板操作 - 薄封装
 * 处理 UI 交互、系统剪贴板读写、调用 Sheet 核心方法
 */

import { generateClipboardHTML, generateClipboardText, parseHTMLStyle } from '../core/Canvas/clipboard.js';
import { fromParsedHTML, fromPlainText, PasteMode, PasteOperation } from '../core/Sheet/Clipboard.js';

/**
 * 复制选中区域
 */
export function copy() {
    const sheet = this.SN.activeSheet;
    const area = sheet.activeAreas[sheet.activeAreas.length - 1];

    // 调用 Sheet 核心方法存储数据
    sheet.copy(area, false);

    // 同步写入系统剪贴板（HTML + 纯文本双格式）
    writeToSystemClipboard(sheet, area);

    return true;
}

/**
 * 剪切选中区域
 */
export function cut() {
    const sheet = this.SN.activeSheet;
    const area = sheet.activeAreas[sheet.activeAreas.length - 1];

    // 调用 Sheet 核心方法存储数据（标记为剪切）
    sheet.cut(area);

    // 同步写入系统剪贴板
    writeToSystemClipboard(sheet, area);

    return true;
}

/**
 * 粘贴（默认全部粘贴）
 * @param {string} mode - 粘贴模式
 */
export async function paste(mode = 'all') {
    const sheet = this.SN.activeSheet;
    const targetCell = sheet.activeCell;

    // 优先使用内部剪贴板
    if (sheet.hasClipboardData()) {
        const options = {
            mode: mode,
            operation: PasteOperation.NONE,
            skipBlanks: false,
            transpose: mode === 'transpose'
        };

        const success = sheet.paste({ r: targetCell.r, c: targetCell.c }, options);
        if (success) {
            // 触发 Volatile 函数重算
            if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
                this.SN.DependencyGraph.recalculateVolatile();
            }
            this.SN.Canvas.r();
        }
        return success;
    }

    // 回退到系统剪贴板
    try {
        const data = await readFromSystemClipboard();
        if (data) {
            const options = {
                mode: mode,
                operation: PasteOperation.NONE,
                skipBlanks: false,
                transpose: mode === 'transpose',
                externalData: data
            };

            const success = sheet.paste({ r: targetCell.r, c: targetCell.c }, options);
            if (success) {
                // 触发 Volatile 函数重算
                if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
                    this.SN.DependencyGraph.recalculateVolatile();
                }
                this.SN.Canvas.r();
            }
            return success;
        }
    } catch (e) {
        console.warn('读取系统剪贴板失败:', e);
    }

    return false;
}

/**
 * 清除剪贴板
 */
export function clearClipboard() {
    this.SN.activeSheet.clearClipboard();
}

// ==================== 内部辅助函数 ====================

/**
 * 写入系统剪贴板
 */
function writeToSystemClipboard(sheet, area) {
    try {
        // 生成 HTML 格式
        const htmlData = generateClipboardHTML(sheet, area);
        // 生成纯文本格式
        const textData = generateClipboardText(sheet, area);

        // 使用 ClipboardItem 同时写入两种格式
        const blob = new ClipboardItem({
            'text/html': new Blob([htmlData], { type: 'text/html' }),
            'text/plain': new Blob([textData], { type: 'text/plain' })
        });

        navigator.clipboard.write([blob]).catch(err => {
            // 降级方案：只写入纯文本
            console.warn('写入 HTML 失败，降级为纯文本:', err);
            navigator.clipboard.writeText(textData);
        });
    } catch (e) {
        console.warn('写入系统剪贴板失败:', e);
    }
}

/**
 * 从系统剪贴板读取数据
 */
async function readFromSystemClipboard() {
    try {
        const items = await navigator.clipboard.read();

        for (const item of items) {
            // 优先尝试 HTML 格式
            if (item.types.includes('text/html')) {
                const blob = await item.getType('text/html');
                const htmlData = await blob.text();
                const parsed = parseHTMLStyle(htmlData);
                if (parsed) {
                    return fromParsedHTML(parsed);
                }
            }

            // 回退到纯文本
            if (item.types.includes('text/plain')) {
                const blob = await item.getType('text/plain');
                const textData = await blob.text();
                return fromPlainText(textData);
            }
        }
    } catch (e) {
        // 某些浏览器可能不支持 clipboard.read()
        // 尝试使用 readText
        try {
            const textData = await navigator.clipboard.readText();
            return fromPlainText(textData);
        } catch (e2) {
            console.warn('读取剪贴板失败:', e2);
        }
    }

    return null;
}

/**
 * 处理粘贴事件（从 DOM 事件获取数据）
 * @param {ClipboardEvent} event - 粘贴事件
 */
export function handlePasteEvent(event) {
    const sheet = this.SN.activeSheet;
    const targetCell = sheet.activeCell;
    const clipboardData = event.clipboardData || window.clipboardData;

    // 优先尝试 HTML 格式（Excel/网页表格数据）
    const htmlData = clipboardData.getData('text/html');
    if (htmlData) {
        const parsed = parseHTMLStyle(htmlData);
        if (parsed && parsed.data && parsed.data.length > 0) {
            const externalData = fromParsedHTML(parsed);
            const success = sheet.paste({ r: targetCell.r, c: targetCell.c }, {
                mode: PasteMode.ALL,
                externalData
            });
            if (success) {
                // 触发 Volatile 函数重算
                if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
                    this.SN.DependencyGraph.recalculateVolatile();
                }
                this.SN.Canvas.r();
            }
            return true;
        }
    }

    // 其次检查纯图片（截图、图片文件等，没有 HTML 数据时才处理）
    const items = clipboardData.items;
    if (items) {
        for (let i = 0; i < items.length; i++) {
            if (items[i].type.indexOf('image') !== -1) {
                const blob = items[i].getAsFile();
                const reader = new FileReader();
                reader.onload = (e) => {
                    const base64 = e.target.result;
                    sheet.Drawing.addImage(base64, {
                        startCell: { r: targetCell.r, c: targetCell.c }
                    });
                    this.SN.Canvas.r();
                };
                reader.readAsDataURL(blob);
                return true;
            }
        }
    }

    // 回退到纯文本
    const textData = clipboardData.getData('text/plain');
    if (textData) {
        const externalData = fromPlainText(textData);
        const success = sheet.paste({ r: targetCell.r, c: targetCell.c }, {
            mode: PasteMode.ALL,
            externalData
        });
        if (success) {
            // 触发 Volatile 函数重算
            if (this.SN.DependencyGraph && this.SN.calcMode !== 'manual') {
                this.SN.DependencyGraph.recalculateVolatile();
            }
            this.SN.Canvas.r();
        }
        return true;
    }

    return false;
}
