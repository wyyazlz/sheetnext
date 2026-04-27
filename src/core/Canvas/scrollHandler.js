// Scrollbar handlers.

function getScrollMetrics(canvas) {
    const sheet = canvas.activeSheet;
    const frozenH = Math.max(0, (sheet.vi?.fh ?? sheet.headHeight) - sheet.headHeight);
    const frozenW = Math.max(0, (sheet.vi?.fw ?? sheet.indexWidth) - sheet.indexWidth);
    const bodyH = Math.max(1, canvas.viewHeight - sheet.headHeight - frozenH);
    const bodyW = Math.max(1, canvas.viewWidth - sheet.indexWidth - frozenW);
    const minTop = sheet.getScrollTop(sheet._frozenRows || 0);
    const minLeft = sheet.getScrollLeft(sheet._frozenCols || 0);
    return {
        minTop,
        minLeft,
        maxTop: Math.max(minTop, sheet.getTotalHeight() - bodyH),
        maxLeft: Math.max(minLeft, sheet.getTotalWidth() - bodyW)
    };
}

function getTranslateValue(styleValue, axis) {
    const pattern = axis === 'y'
        ? /translateY\(([-\d.]+)px\)/
        : /translateX\(([-\d.]+)px\)/;
    return styleValue ? Number(styleValue.match(pattern)?.[1]) || 0 : 0;
}

export function xDown(event) {
    this._stopScrollInertia?.();
    const pageY = event.pageY;
    const initTop = getTranslateValue(this.rollYS.style.transform, 'y');
    const sheet = this.activeSheet;
    const metrics = getScrollMetrics(this);

    const onMouseMove = (e) => {
        const top = initTop + e.pageY - pageY;
        const res = Math.max(0, Math.min(top, this.maxTop));
        const progress = this.maxTop > 0 ? res / this.maxTop : 0;
        const scrollTop = metrics.minTop + progress * (metrics.maxTop - metrics.minTop);
        this.scrollToPixels(sheet.vi.scrollLeft ?? 0, scrollTop, { fast: true });
    };
    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', () => {
        window.removeEventListener('mousemove', onMouseMove);
    }, { once: true });
}

export function yDown(event) {
    this._stopScrollInertia?.();
    const pageX = event.pageX;
    const initLeft = getTranslateValue(this.rollXS.style.transform, 'x');
    const sheet = this.activeSheet;
    const metrics = getScrollMetrics(this);

    const onMouseMove = (e) => {
        const left = initLeft + e.pageX - pageX;
        const res = Math.max(0, Math.min(left, this.maxLeft));
        const progress = this.maxLeft > 0 ? res / this.maxLeft : 0;
        const scrollLeft = metrics.minLeft + progress * (metrics.maxLeft - metrics.minLeft);
        this.scrollToPixels(scrollLeft, sheet.vi.scrollTop ?? 0, { fast: true });
    };
    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', () => {
        window.removeEventListener('mousemove', onMouseMove);
    }, { once: true });
}

export function xTouchDown(event) {
    if (event.touches.length !== 1) return;
    event.preventDefault();
    this._stopScrollInertia?.();

    const touch = event.touches[0];
    const pageY = touch.pageY;
    const initTop = getTranslateValue(this.rollYS.style.transform, 'y');
    const sheet = this.activeSheet;
    const metrics = getScrollMetrics(this);
    const rowTipPrefix = this.SN.t('layoutScrollTipRow');

    this.scrollTip.style.display = 'block';

    const onTouchMove = (e) => {
        if (e.touches.length !== 1) return;
        e.preventDefault();
        const current = e.touches[0];
        const top = initTop + current.pageY - pageY;
        const res = Math.max(0, Math.min(top, this.maxTop));
        const progress = this.maxTop > 0 ? res / this.maxTop : 0;
        const scrollTop = metrics.minTop + progress * (metrics.maxTop - metrics.minTop);
        this.scrollToPixels(sheet.vi.scrollLeft ?? 0, scrollTop, { fast: true });
        const rowIndex = sheet.getRowIndexByScrollTop(scrollTop);
        this.scrollTip.textContent = `${rowTipPrefix} ${rowIndex + 1}`;
    };

    const onTouchEnd = () => {
        window.removeEventListener('touchmove', onTouchMove);
        window.removeEventListener('touchend', onTouchEnd);
        window.removeEventListener('touchcancel', onTouchEnd);
        this.scrollTip.style.display = 'none';
        this.r();
    };

    window.addEventListener('touchmove', onTouchMove, { passive: false });
    window.addEventListener('touchend', onTouchEnd, { once: true });
    window.addEventListener('touchcancel', onTouchEnd, { once: true });
}

export function yTouchDown(event) {
    if (event.touches.length !== 1) return;
    event.preventDefault();
    this._stopScrollInertia?.();

    const touch = event.touches[0];
    const pageX = touch.pageX;
    const initLeft = getTranslateValue(this.rollXS.style.transform, 'x');
    const sheet = this.activeSheet;
    const metrics = getScrollMetrics(this);
    const colTipPrefix = this.SN.t('layoutScrollTipCol');

    this.scrollTip.style.display = 'block';

    const onTouchMove = (e) => {
        if (e.touches.length !== 1) return;
        e.preventDefault();
        const current = e.touches[0];
        const left = initLeft + current.pageX - pageX;
        const res = Math.max(0, Math.min(left, this.maxLeft));
        const progress = this.maxLeft > 0 ? res / this.maxLeft : 0;
        const scrollLeft = metrics.minLeft + progress * (metrics.maxLeft - metrics.minLeft);
        this.scrollToPixels(scrollLeft, sheet.vi.scrollTop ?? 0, { fast: true });
        const colIndex = sheet.getColIndexByScrollLeft(scrollLeft);
        this.scrollTip.textContent = `${colTipPrefix} ${this.Utils.numToChar(colIndex)}`;
    };

    const onTouchEnd = () => {
        window.removeEventListener('touchmove', onTouchMove);
        window.removeEventListener('touchend', onTouchEnd);
        window.removeEventListener('touchcancel', onTouchEnd);
        this.scrollTip.style.display = 'none';
        this.r();
    };

    window.addEventListener('touchmove', onTouchMove, { passive: false });
    window.addEventListener('touchend', onTouchEnd, { once: true });
    window.addEventListener('touchcancel', onTouchEnd, { once: true });
}
