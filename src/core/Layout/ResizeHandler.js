/**
 * 尺寸调整处理模块
 * 负责处理容器尺寸变化和响应式布局
 */

/**
 * 初始化ResizeObserver
 */
export function initResizeObserver() {
    if (typeof ResizeObserver !== 'undefined') {
        this._resizeObserver = new ResizeObserver(entries => {
            this.throttleResize();
        });
        this._resizeObserver.observe(this.SN.containerDom);
    }
}

/**
 * 节流处理resize事件
 */
export function throttleResize() {
    if (this._resizeTimer) {
        clearTimeout(this._resizeTimer);
    }

    this._resizeTimer = setTimeout(() => {
        this.handleAllResize();
        this._resizeTimer = null;
    }, 100);
}

/**
 * 处理所有resize相关的逻辑
 */
export function handleAllResize() {
    const currentContainerWidth = this.SN.containerDom.offsetWidth;
    const currentContainerHeight = this.SN.containerDom.offsetHeight;
    const currentIsSmallWindow = this.isSmallWindow;

    // 检查容器尺寸变化（宽度或高度）
    if (this._lastContainerWidth !== currentContainerWidth ||
        this._lastContainerHeight !== currentContainerHeight) {
        this._lastContainerWidth = currentContainerWidth;
        this._lastContainerHeight = currentContainerHeight;
    }

    // 检查小窗口状态变化
    if (this._lastIsSmallWindow !== currentIsSmallWindow) {
        this._lastIsSmallWindow = currentIsSmallWindow;
        if (currentIsSmallWindow && this._showAIChat && !this._showAIChatWindow) {
            this.showAIChat = false;
        }
    }

    this._syncToolbarModeByWidth?.(currentContainerWidth);
    this._scheduleToolbarAdaptiveLayout?.();

    this.updateCanvasSize();
}

/**
 * 更新Canvas尺寸
 */
export function updateCanvasSize() {
    // 统一处理Canvas视图更新
    this.SN.Canvas?.viewResize();
    this.SN.Canvas?.r();
}
