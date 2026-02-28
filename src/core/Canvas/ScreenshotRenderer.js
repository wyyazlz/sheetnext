// 虚拟截图功能模块

// 虚拟截图功能
// csType: 'topleft' | 'center' - 单元格作为左上角还是中心点
// option: { width, height, compress, hideSelectionFrame } - compress: true(默认0.45) | false(PNG) | 0-1(WebP质量)
async function captureScreenshot(cellAddress, csType = 'topleft', option = {}) {
    const { width = 1600, height = 1000, compress = true, hideSelectionFrame = false } = option

    let sheetName, cellRef;
    if (cellAddress.includes('!')) {
        [sheetName, cellRef] = cellAddress.split('!');
        if (sheetName && sheetName.startsWith("'") && sheetName.endsWith("'")) {
            sheetName = sheetName.slice(1, -1);
        }
    } else {
        cellRef = cellAddress;
        sheetName = null;
    }

    const sheet = this.SN.sheets.find(s => s.name === sheetName) || this.activeSheet
    sheet._init();

    let { r: targetRow, c: targetCol } = this.Utils.cellStrToNum(cellRef);

    // center模式：计算偏移使目标单元格居中，但不超出边界
    if (csType === 'center') {
        const avgColWidth = sheet.defaultColWidth || 80;
        const avgRowHeight = sheet.defaultRowHeight || 25;
        const viewCols = Math.floor(width / avgColWidth / 2);
        const viewRows = Math.floor(height / avgRowHeight / 2);
        targetRow = Math.max(0, targetRow - viewRows);
        targetCol = Math.max(0, targetCol - viewCols);
    }

    // 保存原始状态（直接保存原 vi 引用，避免额外深拷贝）
    const originalVi = sheet.vi;
    const originalViewStart = originalVi ? null : { ...sheet.viewStart };

    try {
        // 临时设置起始位置，使用独立 vi 避免污染原视图对象
        sheet.vi = { rowArr: [targetRow], colArr: [targetCol] };

        // 创建截图 canvas
        if (!this.screenshotCanvas) this.screenshotCanvas = document.createElement('canvas');
        const exportScale = 2;
        const physicalWidth = Math.round(width * exportScale);
        const physicalHeight = Math.round(height * exportScale);
        this.screenshotCanvas.width = physicalWidth;
        this.screenshotCanvas.height = physicalHeight;

        // 调用主渲染方法
        await this.r(null, {
            sheet,
            targetCanvas: this.screenshotCanvas,
            width: physicalWidth,
            height: physicalHeight,
            dpr: exportScale,
            hideSelectionFrame
        });

        // 根据compress参数决定输出格式
        if (compress === false) {
            return this.screenshotCanvas.toDataURL('image/png');
        }
        const quality = typeof compress === 'number' ? compress : 0.45;
        return this.screenshotCanvas.toDataURL('image/webp', quality);
    } finally {
        // 恢复原始状态，避免污染当前视图滚动窗口
        if (originalVi) {
            sheet.vi = originalVi;
        } else if (originalViewStart) {
            sheet.viewStart = originalViewStart;
            sheet._updView();
        }
    }
}

export {
    captureScreenshot
};
