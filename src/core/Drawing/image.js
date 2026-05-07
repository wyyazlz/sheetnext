// 图片初始化：加载图片并自动获取尺寸
export async function _initImage() {
    const img = new Image();
    await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = this._imageBase64;
    });

    // 如果没有设置尺寸，才自动计算（尊重用户手动指定的尺寸）
    if (this._width === null || this._height === null) {
        debugger
        let width, height;

        // 小图片（如图标）不缩放，大图片才应用 DPI 处理
        if (img.naturalWidth > 200 || img.naturalHeight > 200) {
            // 按 Excel 的方式处理：假设图片为 72 DPI，转换为 96 DPI 显示（缩小到 75%）
            width = img.naturalWidth * 0.75;
            height = img.naturalHeight * 0.75;
        } else {
            // 小图原样显示
            width = img.naturalWidth;
            height = img.naturalHeight;
        }

        // 限制最大宽度为 21.59 厘米（816 像素，在 96 DPI 下）
        const MAX_WIDTH = 816;
        if (width > MAX_WIDTH) {
            const ratio = MAX_WIDTH / width;
            width = MAX_WIDTH;
            height *= ratio;
        }

        this._width = Math.round(width);
        this._height = Math.round(height);

        // twoCell 模式：同步计算结束单元格
        if (this.anchorType === 'twoCell') this._syncEndCell();
    }

    // 缓存图片对象
    this._drawCache = img;
    return img;
}

