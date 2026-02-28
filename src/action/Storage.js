/**
 * 本地存储模块
 * 提供浏览器 localStorage 保存/恢复功能
 */

const STORAGE_KEY = 'SN_WORKBOOK';
const MAX_SIZE = 4.5 * 1024 * 1024; // 预留空间，设为 4.5MB

/**
 * 压缩字符串 - 使用 UTF-16 编码压缩
 */
function compress(str) {
    if (!str) return '';
    const bytes = new TextEncoder().encode(str);
    // 两个字节合并为一个 UTF-16 字符
    let result = '';
    for (let i = 0; i < bytes.length; i += 2) {
        const high = bytes[i];
        const low = bytes[i + 1] || 0;
        result += String.fromCharCode((high << 8) | low);
    }
    // 记录原始字节长度（处理奇数长度）
    return bytes.length + '|' + result;
}

/**
 * 解压字符串
 */
function decompress(compressed) {
    if (!compressed) return '';
    try {
        const idx = compressed.indexOf('|');
        if (idx === -1) return null;

        const originalLen = parseInt(compressed.substring(0, idx), 10);
        const data = compressed.substring(idx + 1);

        const bytes = new Uint8Array(originalLen);
        for (let i = 0, j = 0; i < data.length && j < originalLen; i++) {
            const code = data.charCodeAt(i);
            bytes[j++] = code >> 8;
            if (j < originalLen) bytes[j++] = code & 0xff;
        }
        return new TextDecoder().decode(bytes);
    } catch (e) {
        console.error('解压失败:', e);
        return null;
    }
}

/**
 * 获取 localStorage 已用空间
 */
function getStorageUsed() {
    let total = 0;
    for (const key in localStorage) {
        if (localStorage.hasOwnProperty(key)) {
            total += localStorage[key].length * 2;
        }
    }
    return total;
}

/**
 * 格式化字节大小
 */
function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1024 / 1024).toFixed(2) + ' MB';
}

/**
 * 保存到浏览器
 */
export async function saveToLocal() {
    const SN = this.SN;

    try {
        const data = await SN.IO.getData();
        const jsonStr = JSON.stringify(data);
        const originalSize = jsonStr.length * 2;

        // 压缩
        const compressed = compress(jsonStr);
        const compressedSize = compressed.length * 2;
        const ratio = ((1 - compressedSize / originalSize) * 100).toFixed(1);

        // 检查大小
        if (compressedSize > MAX_SIZE) {
            return SN.Utils.toast(SN.t('action.storage.toast.saveTooLarge', {
                originalSize: formatSize(originalSize),
                compressedSize: formatSize(compressedSize),
                limitSize: formatSize(MAX_SIZE)
            }));
        }

        // 检查可用空间
        const used = getStorageUsed();
        const oldData = localStorage.getItem(STORAGE_KEY);
        const oldSize = oldData ? oldData.length * 2 : 0;
        const available = 5 * 1024 * 1024 - used + oldSize;

        if (compressedSize > available) {
            return SN.Utils.toast(SN.t('action.storage.toast.saveInsufficientSpace', {
                compressedSize: formatSize(compressedSize),
                available: formatSize(available)
            }));
        }

        // 保存
        const saveData = JSON.stringify({
            v: 2,
            t: Date.now(),
            n: SN.workbookName,
            d: compressed
        });

        localStorage.setItem(STORAGE_KEY, saveData);
        SN.Utils.toast(SN.t('action.storage.toast.saveSuccess', {
            ratio,
            compressedSize: formatSize(compressedSize)
        }));
    } catch (e) {
        console.error('保存失败:', e);
        if (e.name === 'QuotaExceededError') {
            SN.Utils.toast(SN.t('action.storage.toast.msg001'));
        } else {
            SN.Utils.toast(SN.t('action.storage.toast.saveFailed', { message: e.message }));
        }
    }
}

/**
 * 从浏览器恢复
 */
export async function loadFromLocal() {
    const SN = this.SN;

    try {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (!stored) {
            return SN.Utils.toast(SN.t('action.storage.toast.msg002'));
        }

        const saveData = JSON.parse(stored);
        if (!saveData.d) {
            return SN.Utils.toast(SN.t('action.storage.toast.msg003'));
        }

        // 显示确认对话框
        const time = new Date(saveData.t).toLocaleString();
        const name = saveData.n || '未命名';

        await SN.Utils.modal({
            titleKey: 'action.storage.modal.restore.title',
            title: SN.t('action.storage.modal.restore.title'),
            content: SN.t('action.storage.modal.restore.content', { name, time }),
            confirmKey: 'common.ok',
            confirmText: SN.t('common.ok'),
            cancelKey: 'common.cancel',
            cancelText: SN.t('common.cancel')
        });

        // 解压
        const jsonStr = decompress(saveData.d);
        if (!jsonStr) {
            return SN.Utils.toast(SN.t('action.storage.toast.msg004'));
        }

        const data = JSON.parse(jsonStr);
        const success = SN.IO._setData(data);
        if (success) {
            SN.Utils.toast(SN.t('action.storage.toast.msg005'));
        }
    } catch (e) {
        if (e?.message === 'cancel') return;
        console.error('恢复失败:', e);
        SN.Utils.toast(SN.t('action.storage.toast.restoreFailed', { message: e.message }));
    }
}

/**
 * 检查是否有保存的数据
 */
export function hasLocalData() {
    return !!localStorage.getItem(STORAGE_KEY);
}

/**
 * 清除本地存储
 */
export function clearLocalData() {
    const SN = this.SN;
    localStorage.removeItem(STORAGE_KEY);
    SN.Utils.toast(SN.t('action.storage.toast.msg006'));
}
