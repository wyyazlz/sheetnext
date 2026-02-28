// 辅助函数模块

export function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// 从 NumFmt 模块重新导出
export { modifyDecimalPlaces } from '../core/Cell/NumFmt.js';
