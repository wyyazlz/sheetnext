// 滚动条监听

export function xDown(event) {
    const pageY = event.pageY
    const transformStyle = this.rollYS.style.transform;
    const initTop = transformStyle ? Number(transformStyle.match(/translateY\(([-\d.]+)px\)/)?.[1]) : 0;
    let lastIndex = null
    const sheet = this.activeSheet;
    const rowTipPrefix = this.SN.t('layoutScrollTipRow');
    const frozenRows = sheet._frozenRows || 0;
    // 计算冻结区域的高度
    let frozenHeight = 0;
    for (let i = 0; i < frozenRows; i++) {
        frozenHeight += sheet.getRow(i).height || 0;
    }
    const totalH = sheet.getTotalHeight() - frozenHeight;
    const lastRowH = sheet.getRow(sheet.rowCount - 1)?.height || 0;
    const scrollableH = Math.max(1, totalH - lastRowH);

    const onMouseMove = (e) => {
        const top = initTop + e.pageY - pageY
        let res = Math.max(0, Math.min(top, this.maxTop));
        // 根据滚动条位置计算像素偏移，再找到对应行索引（从冻结行后开始）
        const scrollTop = (res / this.maxTop) * scrollableH;
        let index = sheet.getRowIndexByScrollTop(scrollTop + frozenHeight);
        index = Math.max(index, frozenRows); // 确保不小于冻结行数
        this.rollYS.style.transform = `translateY(${res}px)`
        if (index != lastIndex) {
            lastIndex = index
            sheet.vi.rowArr[0] = index
            this.r();
        }
    };
    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', () => {
        window.removeEventListener('mousemove', onMouseMove);
    }, { once: true });
}

export function yDown(event) {
    const pageX = event.pageX
    const transformStyle = this.rollXS.style.transform;
    const initLeft = transformStyle ? Number(transformStyle.match(/translateX\(([-\d.]+)px\)/)?.[1]) : 0;
    let lastIndex = null
    const sheet = this.activeSheet;
    const frozenCols = sheet._frozenCols || 0;
    // 计算冻结区域的宽度
    let frozenWidth = 0;
    for (let i = 0; i < frozenCols; i++) {
        frozenWidth += sheet.getCol(i)?.width || 0;
    }
    const totalW = sheet.getTotalWidth() - frozenWidth;
    const lastColW = sheet.getCol(sheet.colCount - 1)?.width || 0;
    const scrollableW = Math.max(1, totalW - lastColW);

    const onMouseMove = (e) => {
        const left = initLeft + e.pageX - pageX
        let res = Math.max(0, Math.min(left, this.maxLeft));
        // 根据滚动条位置计算像素偏移，再找到对应列索引（从冻结列后开始）
        const scrollLeft = (res / this.maxLeft) * scrollableW;
        let index = sheet.getColIndexByScrollLeft(scrollLeft + frozenWidth);
        index = Math.max(index, frozenCols); // 确保不小于冻结列数
        this.rollXS.style.transform = `translateX(${res}px)`
        if (index != lastIndex) {
            lastIndex = index
            sheet.vi.colArr[0] = index
            this.r();
        }
    };
    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', () => {
        window.removeEventListener('mousemove', onMouseMove);
    }, { once: true });
}

// 垂直滚动条触摸支持
export function xTouchDown(event) {
    if (event.touches.length !== 1) return;

    // 阻止页面滚动
    event.preventDefault();

    const touch = event.touches[0];
    const pageY = touch.pageY;
    const transformStyle = this.rollYS.style.transform;
    const initTop = transformStyle ? Number(transformStyle.match(/translateY\(([-\d.]+)px\)/)?.[1]) : 0;
    let lastIndex = null;
    const sheet = this.activeSheet;
    const frozenRows = sheet._frozenRows || 0;
    // 计算冻结区域的高度
    let frozenHeight = 0;
    for (let i = 0; i < frozenRows; i++) {
        frozenHeight += sheet.getRow(i)?.height || 0;
    }
    const totalH = sheet.getTotalHeight() - frozenHeight;
    const lastRowH = sheet.getRow(sheet.rowCount - 1)?.height || 0;
    const scrollableH = Math.max(1, totalH - lastRowH);

    // 显示滚动提示
    this.scrollTip.style.display = 'block';

    const onTouchMove = (e) => {
        if (e.touches.length !== 1) return;

        // 阻止页面滚动
        e.preventDefault();

        const touch = e.touches[0];
        const top = initTop + touch.pageY - pageY;
        let res = Math.max(0, Math.min(top, this.maxTop));

        // 根据滚动条位置计算像素偏移，再找到对应行索引（从冻结行后开始）
        const scrollTop = (res / this.maxTop) * scrollableH;
        let index = sheet.getRowIndexByScrollTop(scrollTop + frozenHeight);
        index = Math.max(index, frozenRows); // 确保不小于冻结行数
        this.rollYS.style.transform = `translateY(${res}px)`;
        lastIndex = index;

        // 更新提示内容（位置由CSS居中）
        this.scrollTip.textContent = `${rowTipPrefix} ${index + 1}`;
    };

    const onTouchEnd = () => {
        window.removeEventListener('touchmove', onTouchMove);
        window.removeEventListener('touchend', onTouchEnd);

        // 隐藏提示
        this.scrollTip.style.display = 'none';

        // 只在拖拽结束时渲染一次
        if (lastIndex !== null) {
            sheet.vi.rowArr[0] = lastIndex;
            this.r();
        }
    };

    window.addEventListener('touchmove', onTouchMove, { passive: false });
    window.addEventListener('touchend', onTouchEnd, { once: true });
}

// 水平滚动条触摸支持
export function yTouchDown(event) {
    if (event.touches.length !== 1) return;

    // 阻止页面滚动
    event.preventDefault();

    const touch = event.touches[0];
    const pageX = touch.pageX;
    const transformStyle = this.rollXS.style.transform;
    const initLeft = transformStyle ? Number(transformStyle.match(/translateX\(([-\d.]+)px\)/)?.[1]) : 0;
    let lastIndex = null;
    const sheet = this.activeSheet;
    const colTipPrefix = this.SN.t('layoutScrollTipCol');
    const frozenCols = sheet._frozenCols || 0;
    // 计算冻结区域的宽度
    let frozenWidth = 0;
    for (let i = 0; i < frozenCols; i++) {
        frozenWidth += sheet.getCol(i)?.width || 0;
    }
    const totalW = sheet.getTotalWidth() - frozenWidth;
    const lastColW = sheet.getCol(sheet.colCount - 1)?.width || 0;
    const scrollableW = Math.max(1, totalW - lastColW);

    // 显示滚动提示
    this.scrollTip.style.display = 'block';

    const onTouchMove = (e) => {
        if (e.touches.length !== 1) return;

        // 阻止页面滚动
        e.preventDefault();

        const touch = e.touches[0];
        const left = initLeft + touch.pageX - pageX;
        let res = Math.max(0, Math.min(left, this.maxLeft));

        // 根据滚动条位置计算像素偏移，再找到对应列索引（从冻结列后开始）
        const scrollLeft = (res / this.maxLeft) * scrollableW;
        let index = sheet.getColIndexByScrollLeft(scrollLeft + frozenWidth);
        index = Math.max(index, frozenCols); // 确保不小于冻结列数
        this.rollXS.style.transform = `translateX(${res}px)`;
        lastIndex = index;

        // 更新提示内容（位置由CSS居中）
        this.scrollTip.textContent = `${colTipPrefix} ${this.Utils.numToChar(index)}`;
    };

    const onTouchEnd = () => {
        window.removeEventListener('touchmove', onTouchMove);
        window.removeEventListener('touchend', onTouchEnd);

        // 隐藏提示
        this.scrollTip.style.display = 'none';

        // 只在拖拽结束时渲染一次
        if (lastIndex !== null) {
            sheet.vi.colArr[0] = lastIndex;
            this.r();
        }
    };

    window.addEventListener('touchmove', onTouchMove, { passive: false });
    window.addEventListener('touchend', onTouchEnd, { once: true });
}
