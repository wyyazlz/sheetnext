/**
 * Sizing Processing Module
 * Responsible for handling container size changes and responsive layout
 */

/**
 * Sizing Processing Module
 * Responsible for handling container size changes and responsive layout
 */

/**
 * Initialize ResizeObserver
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
 * Throttling resize event
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
 * Handle all resize related logic
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

    if (!this.SN.containerDom.isConnected || currentContainerWidth <= 0 || currentContainerHeight <= 0) {
        return;
    }

    this._syncToolbarModeByWidth?.(currentContainerWidth);
    this._scheduleToolbarAdaptiveLayout?.();

    this.updateCanvasSize();
}

/**
 * Update Canvas Dimensions
 */
export function updateCanvasSize() {
    // 统一处理Canvas视图更新
    const hasViewport = this.SN.Canvas?.viewResize();
    if (hasViewport === false) return;
    this.SN.Canvas?.r();
}
