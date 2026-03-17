// 导入渲染模块
import * as CellRenderer from './CellRenderer.js';
import * as SheetRenderer from './SheetRenderer.js';
import * as ScreenshotRenderer from './ScreenshotRenderer.js';
import * as IndexRenderer from './IndexRenderer.js';
import * as InputManager from './InputManager.js';
import * as EventManager from './EventManager.js';
import * as SparklineRenderer from './SparklineRenderer.js';
import * as CommentRenderer from './CommentRenderer.js';
import * as CFRenderer from './CFRenderer.js';
import * as FilterRenderer from './FilterRenderer.js';
import * as PivotTableRenderer from './PivotTableRenderer.js';
import FilterPanel from '../AutoFilter/FilterPanel.js';
import LazyDOM from '../Layout/LazyDOM.js';

const INPUT_VISIBLE_Z_INDEX = '999';
const INPUT_HIDDEN_Z_INDEX = '-1';

const _clm = new WeakMap();
export const _cgl = (canvas) => _clm.get(canvas);

/**
 * Canvas Rendering and Interaction Management
 * @title 🖱️ Binding Actions
 * @class
 */
export default class Canvas {
    /** @type {number} */
    dpr;
    /** @type {CanvasRenderingContext2D|null} */
    ctx;
    /** @type {CanvasRenderingContext2D|null} */
    ctxKZ;
    /** @type {CanvasRenderingContext2D|null} */
    HDctx;

    /**
     * @param {Object} SN - SheetNext Main Instance
     * @param {Object} license
     */
    constructor(SN, license) {
        /**
         * SheetNext Main Instance
         * @type {Object}
         */
        this.SN = SN
        _clm.set(this, license);
        /**
         * Tool Method Collection
         * @type {Utils}
         */
        this.Utils = SN.Utils
        /**
         * Canvas Container
         * @type {HTMLElement}
         */
        this.dom = SN.containerDom.querySelector('.sn-canvas')

        /**
         * Show Layer Canvas
         * @type {HTMLCanvasElement}
         */
        this.showLayer = SN.containerDom.querySelector(".sn-show-layer")

        /**
         * Operating Layer Canvas
         * @type {HTMLCanvasElement}
         */
        this.handleLayer = SN.containerDom.querySelector(".sn-handle-layer")
        // 缓冲画布
        /**
         * Buffer Canvas
         * @type {HTMLCanvasElement}
         */
        this.buffer = document.createElement('canvas');
        /**
         * Snapshot Canvas
         * @type {HTMLCanvasElement}
         */
        this.showLayerKZ = document.createElement('canvas');
        /**
         * Brush preview canvas
         * @type {HTMLCanvasElement}
         */
        this.pCanvas = document.createElement('canvas');
        this.pCanvas.width = this.pCanvas.height = 8;
        /**
         * Brush Preview Context
         * @type {CanvasRenderingContext2D}
         */
        this.pCtx = this.pCanvas.getContext('2d');
        // 其他画布信息
        /**
         * Current Active Border Information
         * @type {Object|null}
         */
        this.activeBorderInfo = null;
        /**
         * Multi-select border information
         * @type {Object}
         */
        this.mpBorder = {}
        /**
         * Last padding information
         * @type {Object|null}
         */
        this.lastPadding = null;
        // 鼠标事件
        /**
         * Mouse pointer is currently pointing to an area
         * @type {string|null}
         */
        this.moveAreaName = null; // 鼠标指针当前指向区域 row-col-resize drawing fill-handle active-border
        /**
         * Whether to disable mobile detection
         * @type {boolean}
         */
        this.disableMoveCheck = false;
        /**
         * Whether to disable scrolling
         * @type {boolean}
         */
        this.disableScroll = false; // 是否禁用滚动功能
        // 必要DOM（初始化时获取）
        /**
         * Input Editor
         * @type {HTMLElement}
         */
        this.input = SN.containerDom.querySelector('.sn-input-editor'); // 画布富文本输入框
        /**
         * Formula Bar Input Box
         * @type {HTMLElement}
         */
        this.formulaBar = SN.containerDom.querySelector('.sn-formula-bar'); //工具栏输入框
        this.inputEditing = false;
        /**
         * Area input box
         * @type {HTMLInputElement}
         */
        this.areaInput = SN.containerDom.querySelector('.sn-area-input');
        /**
         * Area Selection Container
         * @type {HTMLElement}
         */
        this.areaBox = SN.containerDom.querySelector('.sn-area-box');
        /**
         * Area Dropdown Container
         * @type {HTMLElement}
         */
        this.areaDropdown = SN.containerDom.querySelector('.sn-area-dropdown');
        /**
         * Landscape scrollbar
         * @type {HTMLElement}
         */
        this.rollX = SN.containerDom.querySelector('.sn-roll-x')
        /**
         * Landscape Scrollbar Slider
         * @type {HTMLElement}
         */
        this.rollXS = SN.containerDom.querySelector('.sn-roll-x>div')
        /**
         * Portrait scrollbar
         * @type {HTMLElement}
         */
        this.rollY = SN.containerDom.querySelector('.sn-roll-y')
        /**
         * Portrait Scrollbar Slider
         * @type {HTMLElement}
         */
        this.rollYS = SN.containerDom.querySelector('.sn-roll-y>div')
        /**
         * Graphics Container
         * @type {HTMLElement}
         */
        this.drawingsCon = SN.containerDom.querySelector('.sn-drawings-con')
        /**
         * Slicer Container
         * @type {HTMLElement}
         */
        this.slicersCon = SN.containerDom.querySelector('.sn-slicers-con')

        /**
         * Maximum longitudinal scrolling distance
         * @type {number}
         */
        this.maxTop = 0;
        /**
         * Maximum lateral scrolling distance
         * @type {number}
         */
        this.maxLeft = 0;

        /**
         * Lazy loading Dom manager
         * @type {LazyDOM}
         */
        this.lazyDOM = new LazyDOM(SN);

        /**
         * Filter Panel Manager
         * @type {FilterPanel}
         */
        this.filterPanelManager = new FilterPanel(this);

        /**
         * Render Count
         * @type {number}
         */
        this.rCount = 0

        /**
         * Current hover comment
         * @type {Object|null}
         */
        this.hoveredComment = null;

        /**
         * Statistical Sum: Dom
         * @type {HTMLElement|null}
         */
        this.statSum = SN.containerDom.querySelector('[data-stat="sum"]');
        /**
         * Statistics Count: Dom
         * @type {HTMLElement|null}
         */
        this.statCount = SN.containerDom.querySelector('[data-stat="count"]');
        /**
         * Statistical average: Dom
         * @type {HTMLElement|null}
         */
        this.statAvg = SN.containerDom.querySelector('[data-stat="avg"]');
        this._statRafId = null;

        /**
         * Highlight color, null means closed
         * @type {string|null}
         */
        this.highlightColor = null; // 高亮颜色，null 表示关闭
        // 护眼模式
        /**
         * Whether to turn on eye protection mode
         * @type {boolean}
         */
        this.eyeProtectionMode = false;
        /**
         * Eye protection color
         * @type {string}
         */
        this.eyeProtectionColor = 'rgb(204,232,207)';
        //事件监听
        this.addEventListeners();
        this._initAreaSelector();
        this.viewResize()
    }


    /**
     * Padding Type
     * @type {string}     */
    /**
     * Padding Type
     * @type {string}     */

    get paddingType() { return this.lazyDOM.paddingType; }
    /**
     * Hyperlink Jump Button
     * @type {HTMLElement|null}     */
    /**
     * Hyperlink Jump Button
     * @type {HTMLElement|null}     */

    get hyperlinkJumpBtn() { return this.lazyDOM.hyperlinkBtn; }
    /**
     * Hyperlink Tips
     * @type {HTMLElement|null}     */
    /**
     * Hyperlink Tips
     * @type {HTMLElement|null}     */

    get hyperlinkTip() { return this.lazyDOM.hyperlinkTip; }
    /**
     * Data Validation Selector
     * @type {HTMLElement|null}     */
    /**
     * Data Validation Selector
     * @type {HTMLElement|null}     */

    get validSelect() { return this.lazyDOM.validSelect; }
    /**
     * Tips for data validation
     * @type {HTMLElement|null}     */
    /**
     * Tips for data validation
     * @type {HTMLElement|null}     */

    get validTip() { return this.lazyDOM.validTip; }
    /**
     * Scroll Tips
     * @type {HTMLElement|null}     */
    /**
     * Scroll Tips
     * @type {HTMLElement|null}     */

    get scrollTip() { return this.lazyDOM.scrollTip; }

    /**
     * Re-fetch dom length after updating canvas zoom/length-width
     * @returns {boolean}
     */
    viewResize() {
        this.dpr = window.devicePixelRatio || 1;
        const { clientWidth: w, clientHeight: h } = this.dom;
        const hasViewport = this.dom.isConnected && w > 0 && h > 0;
        [this.showLayer, this.handleLayer, this.buffer, this.showLayerKZ].forEach(canvas => {
            const canvasWidth = hasViewport ? w * this.dpr : 0;
            const canvasHeight = hasViewport ? h * this.dpr : 0;
            canvas.width = canvasWidth;
            canvas.height = canvasHeight;
            canvas.style.width = (hasViewport ? w : 0) + 'px'
            canvas.style.height = (hasViewport ? h : 0) + 'px'
        })

        this._width = null
        this._height = null

        this.ctx = this.showLayer.getContext('2d');
        this.ctxKZ = this.showLayerKZ.getContext('2d'); // 快照1
        this.HDctx = this.handleLayer.getContext('2d');

        if (!hasViewport) return false;

        this.#initBufferContext();

        // 窗口变化时更新滚动条尺寸
        if (this.activeSheet) {
            this.updateScrollBarSize();
            this.updateOverlayContainers();
        }
        return true;
    }

    /**
     * Visible Width
     * @type {number}     */
    /**
     * Visible Width
     * @type {number}     */

    get width() {
        return this._width ?? (this._width = this.showLayer.offsetWidth)
    }

    /**
     * Visible Height
     * @type {number}     */
    /**
     * Visible Height
     * @type {number}     */

    get height() {
        return this._height ?? (this._height = this.showLayer.offsetHeight)
    }

    /**
     * Logical View Width
     * @type {number}     */
    /**
     * Logical View Width
     * @type {number}     */

    get viewWidth() {
        return this.width / (this.activeSheet?.zoom ?? 1);
    }

    /**
     * Logical View Height
     * @type {number}     */
    /**
     * Logical View Height
     * @type {number}     */

    get viewHeight() {
        return this.height / (this.activeSheet?.zoom ?? 1);
    }

    /**
     * Physical pixel to logical coordinates
     * @param {number} value - Physical pixels
     * @returns {number}
     */
    toLogical(value) {
        return value / (this.activeSheet?.zoom ?? 1);
    }

    /**
     * Logical Coordinates to Physical Pixels
     * @param {number} value - Logical coordinates
     * @returns {number}
     */
    toPhysical(value) {
        return value * (this.activeSheet?.zoom ?? 1);
    }

    /**
     * Pixel Ratio
     * @type {number}     */
    /**
     * Pixel Ratio
     * @type {number}     */

    get pixelRatio() {
        return this._pixelRatio ?? ((this.dpr || 1) * (this.activeSheet?.zoom ?? 1));
    }

    /**
     * Half Pixel Offset
     * @type {number}     */
    /**
     * Half Pixel Offset
     * @type {number}     */

    get halfPixel() {
        return 0.5 / this.pixelRatio;
    }

    /**
     * Convert to value in pixel ratio
     * @param {number} value - Raw Value
     * @returns {number}
     */
    px(value) {
        return value / this.pixelRatio;
    }

    /**
     * Alignment Pixel Ratio
     * @param {number} value - Raw Value
     * @returns {number}
     */
    snap(value) {
        const ratio = this.pixelRatio;
        return Math.round(value * ratio) / ratio;
    }

    /**
     * Align Rectangle to Pixel Ratio
     * @param {number} x - X
     * @param {number} y - Y
     * @param {number} w - Wide
     * @param {number} h - High
     * @returns {{x: number, y: number, w: number, h: number}}
     */
    snapRect(x, y, w, h) {
        const ratio = this.pixelRatio;
        const x1 = Math.round(x * ratio) / ratio;
        const y1 = Math.round(y * ratio) / ratio;
        const x2 = Math.round((x + w) * ratio) / ratio;
        const y2 = Math.round((y + h) * ratio) / ratio;
        return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
    }

    /**
     * Current Active Sheet
     * @type {Sheet}     */
    /**
     * Current Active Sheet
     * @type {Sheet}     */

    get activeSheet() {
        return this.SN.activeSheet
    }

    /**
     * Whether or not in input edit state
     * @type {boolean}
     */
    get inputEditing() {
        return this._inputEditing === true;
    }

    set inputEditing(value) {
        const nextValue = Boolean(value);
        if (this._inputEditing === nextValue) return;
        this._inputEditing = nextValue;
        this.#syncInputEditorState();
    }

    #syncInputEditorState() {
        if (!this.input) return;
        if (this._inputEditing) {
            this.input.style.zIndex = INPUT_VISIBLE_Z_INDEX;
            this.input.style.opacity = '1';
            this.input.style.pointerEvents = 'auto';
            return;
        }
        this.input.style.zIndex = INPUT_HIDDEN_Z_INDEX;
        this.input.style.opacity = '0';
        this.input.style.pointerEvents = 'none';
    }

    /**
     * Gets the position of the event in the canvas (logical coordinates)
     * @ param {MouseEvent | {clientX: number, clientY: number}} event - Mouse event or object with coordinates
     * @returns {{x: number, y: number}}
     */
    getEventPosition(event) {
        const rect = this.handleLayer.getBoundingClientRect();
        const x = this.toLogical(event.clientX - rect.left);
        const y = this.toLogical(event.clientY - rect.top);
        return { x, y };
    }

    /**
     * Get cell index by location
     * @param {number} x - Logical coordinates X
     * @param {number} y - Logical coordinates Y
     * @returns {{c: number, r: number}}
     */
    getCellIndexByPosition(x, y) {
        let cellX = this.activeSheet.indexWidth;
        let cellY = this.activeSheet.headHeight;
        let clickedRow = -1;

        this.activeSheet.vi.rowsIndex.forEach(rowIndex => {
            const rowHeight = this.activeSheet.getRow(rowIndex).height;
            if (y >= cellY && y < cellY + rowHeight && clickedRow === -1) {
                clickedRow = rowIndex;
            }
            cellY += rowHeight;
        });

        let clickedCol = -1;
        cellX = this.activeSheet.indexWidth;
        this.activeSheet.vi.colsIndex.forEach(colIndex => {
            const colWidth = this.activeSheet.getCol(colIndex).width;
            if (x >= cellX && x < cellX + colWidth && clickedCol === -1) {
                clickedCol = colIndex;
            }
            cellX += colWidth;
        });

        if (clickedCol == -1 && x >= this.activeSheet.vi.w) clickedCol = this.activeSheet.vi.colsIndex[this.activeSheet.vi.colsIndex.length - 1]
        if (clickedRow == -1 && y >= this.activeSheet.vi.h) clickedRow = this.activeSheet.vi.rowsIndex[this.activeSheet.vi.rowsIndex.length - 1];

        return { c: clickedCol, r: clickedRow };
    }

    /**
     * Apply zoom and synchronize scrollbar and input box positions
     * @returns {void}
     */
    applyZoom() {
        if (!this.bc) return;
        const zoom = this.activeSheet?.zoom ?? 1;
        const scale = (this.dpr || 1) * zoom;
        this._pixelRatio = scale;
        this.bc.setTransform(scale, 0, 0, scale, 0, 0);
        this.bc.translate(-0.5 / scale, -0.5 / scale);
        this._width = null;
        this._height = null;
        if (this.activeSheet) {
            this.updateScrollBarSize();
            this.updateOverlayContainers();
            this.updateScrollBar();
        }
        if (this.inputEditing && this.currentCell) {
            this.updateInputPosition?.();
        }
    }

    /**
     * Clear Buffered Canvas (Physical Pixels)
     * @param {number} [width=this.buffer.width] - Width
     * @param {number} [height=this.buffer.height] - Height
     * @returns {void}
     */
    clearBuffer(width = this.buffer.width, height = this.buffer.height) {
        if (!this.bc) return;
        this.bc.save();
        this.bc.setTransform(1, 0, 0, 1, 0, 0);
        this.bc.clearRect(0, 0, width, height);
        this.bc.restore();
    }

    // 初始化缓冲画布上下文
    #initBufferContext() {
        this.bc = this.buffer.getContext('2d');
        const zoom = this.activeSheet?.zoom ?? 1;
        const scale = (this.dpr || 1) * zoom;
        this._pixelRatio = scale;
        this.bc.setTransform(scale, 0, 0, scale, 0, 0);
        this.bc.font = '14.6px 宋体';
        this.bc.lineWidth = 1;
        this.bc.translate(-0.5 / scale, -0.5 / scale);
    }

    // 更新滚动条尺寸（窗口变化或切换sheet时调用）
    /**
     * Update scrollbar dimensions
     * @returns {void}
     */
    updateScrollBarSize() {
        const sheet = this.activeSheet;
        if (!sheet) return;

        const viewH = Math.max(0, this.viewHeight - sheet.headHeight);
        const viewW = Math.max(0, this.viewWidth - sheet.indexWidth);
        const totalH = sheet.getTotalHeight();
        const totalW = sheet.getTotalWidth();

        const trackH = this.rollY.offsetHeight;
        const trackW = this.rollX.offsetWidth;


        const yRatio = Math.min(0.3, viewH / totalH);
        const xRatio = Math.min(0.3, viewW / totalW);

        const scrollBarH = Math.max(30, trackH * yRatio);
        const scrollBarW = Math.max(30, trackW * xRatio);

        this.rollYS.style.height = scrollBarH + 'px';
        this.rollXS.style.width = scrollBarW + 'px';

        this.maxTop = trackH - scrollBarH;
        this.maxLeft = trackW - scrollBarW;
    }


    /**
     * Update Overlay Container Dimensions
     * @returns {void}
     */
    updateOverlayContainers() {
        const sheet = this.activeSheet;
        if (!sheet) return;

        const zoom = sheet.zoom ?? 1;
        const left = sheet.indexWidth * zoom;
        const top = sheet.headHeight * zoom;
        const width = Math.max(0, this.width - left);
        const height = Math.max(0, this.height - top);

        const applyBounds = (el) => {
            if (!el) return;
            el.style.left = `${left}px`;
            el.style.top = `${top}px`;
            el.style.width = `${width}px`;
            el.style.height = `${height}px`;
        };

        applyBounds(this.slicersCon);
        applyBounds(this.drawingsCon);
    }


    /**
     * Update scrollbar position
     * @returns {void}
     */
    updateScrollBar() {
        const sheet = this.activeSheet;
        if (!sheet || !sheet.vi) return;

        const totalH = sheet.getTotalHeight();
        const totalW = sheet.getTotalWidth();

        const lastRowH = sheet.getRow(sheet.rowCount - 1)?.height || 0;
        const lastColW = sheet.getCol(sheet.colCount - 1)?.width || 0;


        const scrollableH = Math.max(0, totalH - lastRowH);
        const scrollableW = Math.max(0, totalW - lastColW);

        // 当前滚动位置
        const scrollTop = sheet.getScrollTop(sheet.vi.rowArr[0]);
        const scrollLeft = sheet.getScrollLeft(sheet.vi.colArr[0]);


        const top = scrollableH > 0 ? (scrollTop / scrollableH) * this.maxTop : 0;
        const left = scrollableW > 0 ? (scrollLeft / scrollableW) * this.maxLeft : 0;

        this.rollYS.style.transform = `translateY(${Math.min(top, this.maxTop)}px)`;
        this.rollXS.style.transform = `translateX(${Math.min(left, this.maxLeft)}px)`;
    }



    #rMpBorder(sheet) {
        if (Object.keys(this.mpBorder).length < 1) return
        const { x, y, w, h, inView } = sheet.getAreaInviewInfo(this.mpBorder)
        if (!inView) return
        const rect = this.snapRect(x, y, w, h);
        const dash = this.px(5);
        this.bc.beginPath();
        this.bc.lineWidth = this.px(1.5);
        this.bc.setLineDash([dash, dash]); // 设置虚线样式,数字表示线段和间隙的长度
        this.bc.strokeStyle = '#333';
        this.bc.strokeRect(rect.x, rect.y, rect.w, rect.h);
        this.bc.closePath();
        this.bc.lineWidth = this.px(1);
        this.bc.setLineDash([]);
    }

    #rFreezeLine(sheet) {
        this.bc.beginPath();
        this.bc.strokeStyle = "#36c"
        this.bc.lineWidth = this.px(1);
        const vi = sheet.vi
        if (vi.frozenCols.length != 0) {
            const fx = this.snap(vi.fw);
            this.bc.moveTo(fx, 0);
            this.bc.lineTo(fx, this.viewHeight);
            this.bc.stroke();
        }
        if (vi.frozenRows.length != 0) {
            const fy = this.snap(vi.fh);
            this.bc.moveTo(0, fy);
            this.bc.lineTo(this.viewWidth, fy);
            this.bc.stroke();
        }
    }

    /**
     * Get the cell index corresponding to the event
     * @ param {MouseEvent | {clientX: number, clientY: number}} event - Mouse event or object with coordinates
     * @returns {{r:number, c:number}}
     */
    getCellIndex(event) {
        const { x, y } = this.getEventPosition(event);
        return this.getCellIndexByPosition(x, y);
    }

    // 更新统计栏（均值、计数、求和）
    _updateStats() {
        if (!this.statSum || !this.statCount || !this.statAvg) return;
        if (this._statRafId) cancelAnimationFrame(this._statRafId);

        this._statRafId = requestAnimationFrame(() => {
            const sheet = this.activeSheet;
            let sum = 0, count = 0;
            const sumLabel = this.SN.t('layout.zoomSum');
            const countLabel = this.SN.t('layout.zoomCount');
            const avgLabel = this.SN.t('layout.zoomAverage');

            sheet.eachCells(sheet.activeAreas, (r, c) => {
                const cell = sheet.getCell(r, c);
                if (cell.isMerged && cell.master) return;
                const val = cell._calcVal;
                if (typeof val === 'number' && !isNaN(val)) {
                    sum += val;
                    count++;
                }
            });

            this.statSum.textContent = `${sumLabel}: ${count ? sum.toLocaleString() : '-'}`;
            this.statCount.textContent = `${countLabel}: ${count || '-'}`;
            this.statAvg.textContent = `${avgLabel}: ${count ? (sum / count).toLocaleString() : '-'}`;
        });
    }

    /**
     * Set highlight color and sync toolbar status
     * @param {MouseEvent} event - Trigger event
     * @returns {void}
     */
    setHighlight(event) {
        const btn = event.target.closest('.sn-hl-btn');
        if (!btn) return;

        this.highlightColor = btn.dataset.color || null;
        const color = this.highlightColor;
        const panel = this.SN.containerDom.querySelector('.sn-highlight-menu .sn-highlight-panel');
        if (panel) {
            panel.querySelectorAll('.sn-hl-btn').forEach(b => {
                b.classList.toggle('sn-btn-selected', b.dataset.color === (color ?? ''));
            });
            const dropdown = panel.closest('.sn-dropdown');
            if (dropdown) dropdown.classList.toggle('sn-hl-active', !!color);
        }

        this.r();
    }

    /**
     * Toggle eye protection mode
     * @returns {void}
     */
    toggleEyeProtection() {
        this.eyeProtectionMode = !this.eyeProtectionMode;


        const btn = this.SN.containerDom.querySelector('[data-icon="huyan_"]')?.closest('.sn-tools-item');
        if (btn) btn.classList.toggle('sn-btn-selected', this.eyeProtectionMode);

        this.r();
    }



    /**
     * Render Canvas (Snapshot mode supported)
     * @param {string} [type] - Rendering type, 's' for snapshot rendering
     * @param {Object} [options] - Rendering Options
     * @param {HTMLCanvasElement} [options.targetCanvas] - Target Canvas (Screenshot Mode)
     * @param {number} [options.width] - Screenshot Width
     * @param {number} [options.height] - Screenshot height
     * @param {number} [options.dpr] - Specify dpr
     * @param {Sheet} [options.sheet] - Assign Sheets
     * @returns {Promise<void>}
     */
    async r(type, options = {}) {
        const _startTime = performance.now();

        const isScreenshot = !!options.targetCanvas;
        const originalDpr = this.dpr;

        if (isScreenshot && Number.isFinite(options.dpr)) {
            this.dpr = options.dpr;
        }

        let sheet = this.activeSheet;

        try {
            const hasViewport = isScreenshot
                ? options.width > 0 && options.height > 0
                : this.dom.isConnected && this.showLayer.width > 0 && this.showLayer.height > 0 && this.buffer.width > 0 && this.buffer.height > 0;
            if (!sheet || !hasViewport) return;
            if (!isScreenshot) {
                // 节流非截图模式的渲染频率（rAF + 脏标记，保证不丢最终帧）
                if (this._rThrottleLock) {
                    this._rDirty = true;
                    this._rDirtyType = type;
                    return;
                }
                this._rThrottleLock = true;
                this._rDirty = false;
                requestAnimationFrame(() => {
                    this._rThrottleLock = false;
                    if (this._rDirty) {
                        this._rDirty = false;
                        this.r(this._rDirtyType);
                    }
                });
                this.ctx.clearRect(0, 0, this.showLayer.width, this.showLayer.height); // 清除显示画布
            } else {

                sheet = options.sheet || this.activeSheet;
                // 临时调整 buffer 尺寸
                options._originalBufferWidth = this.buffer.width;
                options._originalBufferHeight = this.buffer.height;
                this.buffer.width = options.width;
                this.buffer.height = options.height;
                this.#initBufferContext();

                options._originalWidth = this._width;
                options._originalHeight = this._height;
                this._width = options.width / this.dpr;
                this._height = options.height / this.dpr;
            }


            this.clearBuffer(isScreenshot ? options.width : undefined, isScreenshot ? options.height : undefined);

            if (type == 's') {
                this.ctx.drawImage(this.showLayerKZ, 0, 0);
                this.rIndex(sheet);
                this._updateStats();
                return
            }

            // 更新视图信息（截图模式传入宽高）
            if (isScreenshot) {
                const dpr = this.dpr || 1;
                // updView 需要逻辑像素
                sheet._updView(options.width / dpr, options.height / dpr);
            } else {
                sheet._updView();
                this.updateOverlayContainers();
            }

            this.rBaseData(sheet) // 绘制基础数据

            this.rConditionalFormats(sheet) // 绘制条件格式

            this.rMerged(sheet) // 绘制合并的单元格

            this.rSparklines(sheet)

            this.rCommentIndicators(sheet)

            this.rCommentBoxes(sheet)

            this.rFilterIcons(sheet)

            this.rPivotTablePlaceholders(sheet)

            this.rBorder(this.borders)

            this.#rMpBorder(sheet)

            this.#rFreezeLine(sheet)


            if (!isScreenshot) {
                this.ctxKZ.clearRect(0, 0, this.showLayer.width, this.showLayer.height);
                this.ctxKZ.drawImage(this.buffer, 0, 0);
            }


            if (isScreenshot) {
                this.rIndex(sheet, true, { hideSelectionFrame: !!options.hideSelectionFrame });
                const nowRCount = ++this.rCount;
                await this.rDrawing(nowRCount, options);

                const lightInfo = this.getLight();
                const headerLightInfo = options.hideSelectionFrame ? { rows: [], columns: [] } : lightInfo;
                this.rRowColHeaders(sheet, headerLightInfo);
                const targetCtx = options.targetCanvas.getContext('2d');
                targetCtx.clearRect(0, 0, options.width, options.height);
                targetCtx.drawImage(this.buffer, 0, 0);
                // 恢复 buffer 尺寸
                if (this.dpr !== originalDpr) {
                    this.dpr = originalDpr;
                }
                this.buffer.width = options._originalBufferWidth;
                this.buffer.height = options._originalBufferHeight;
                this.#initBufferContext();
                // 恢复宽高
                this._width = options._originalWidth;
                this._height = options._originalHeight;
            } else {
                this.rIndex(sheet);
                // 更新滚动条尺寸和位置
                this.updateScrollBarSize();
                this.updateScrollBar();
            }

            // -------------------- 平均3毫秒每次渲染 --------------------

        } finally {
            if (isScreenshot && this.dpr !== originalDpr) {
                this.dpr = originalDpr;
                this._pixelRatio = null;
            }
            // console.log(`Render time: ${(performance.now() - _startTime).toFixed(2)}ms`);
        }

    }

}

// 使用原型链挂载渲染与交互模块方法
Object.assign(Canvas.prototype, {
    // 单元格渲染
    rCell: CellRenderer.rCell,
    rCellBackground: CellRenderer.rCellBackground,
    rCellText: CellRenderer.rCellText,
    rCellTextWithOverflow: CellRenderer.rCellTextWithOverflow,
    rMerged: CellRenderer.rMerged,
    rBorder: CellRenderer.rBorder,
    rBaseData: CellRenderer.rBaseData,

    // 工作表渲染
    getOffset: SheetRenderer.getOffset,
    rDrawing: SheetRenderer.rDrawing,

    // 截图渲染
    captureScreenshot: ScreenshotRenderer.captureScreenshot,

    // 行列标渲染
    getLight: IndexRenderer.getLight,
    rAllBtn: IndexRenderer.rAllBtn,
    rRowColHeaders: IndexRenderer.rRowColHeaders,
    rHighlightBars: IndexRenderer.rHighlightBars,
    rPageBreaks: IndexRenderer.rPageBreaks,
    rIndex: IndexRenderer.rIndex,
    rDom: IndexRenderer.rDom,
    _updateContextualToolbar: IndexRenderer._updateContextualToolbar,

    // 输入管理
    showCellInput: InputManager.showCellInput,
    updateInputPosition: InputManager.updateInputPosition,
    updInputValue: InputManager.updInputValue,

    // 事件管理
    addEventListeners: EventManager.addEventListeners,

    // 迷你图渲染
    rSparklines: SparklineRenderer.rSparklines,

    // 批注渲染
    rCommentIndicators: CommentRenderer.rCommentIndicators,
    rCommentBoxes: CommentRenderer.rCommentBoxes,

    // 条件格式渲染
    rConditionalFormats: CFRenderer.rConditionalFormats,
    rDataBar: CFRenderer.rDataBar,

    // 筛选图标渲染
    rFilterIcons: FilterRenderer.rFilterIcons,
    clearFilterIconPositions: FilterRenderer.clearFilterIconPositions,
    getFilterIconAtPosition: FilterRenderer.getFilterIconAtPosition,
    drawFilterIcon: FilterRenderer.drawFilterIcon,

    // 透视表占位渲染
    rPivotTablePlaceholders: PivotTableRenderer.rPivotTablePlaceholders,
    getPivotTableAtPosition: PivotTableRenderer.getPivotTableAtPosition
});

// 区域选择器初始化
Canvas.prototype._initAreaSelector = function() {
    const { areaBox, areaDropdown, areaInput, SN } = this;
    if (!areaBox || !areaDropdown || !areaInput) return;

    const closeDropdown = () => areaBox.classList.remove('open');

    // 解析区域字符串为数字对象
    const parseRange = (rangeStr) => {
        const clean = rangeStr.replace(/\$/g, '');
        const parts = clean.split(':');
        const s = SN.Utils.cellStrToNum(parts[0]);
        const e = parts[1] ? SN.Utils.cellStrToNum(parts[1]) : { ...s };
        return s && e ? { s, e } : null;
    };


    const gotoRef = (ref) => {
        if (!ref) return;
        const match = ref.match(/^(?:(.+?)!)?(.+)$/);
        if (!match) return;
        const [, sheetName, range] = match;

        // 确定目标 sheet
        let targetSheet = SN.activeSheet;
        if (sheetName) {
            const cleanName = sheetName.replace(/^'+|'+$/g, '');
            targetSheet = SN.sheets.find(s => s.name === cleanName);
            if (!targetSheet) return;
        }

        // 解析区域
        const rangeObj = parseRange(range);
        if (!rangeObj) return;


        targetSheet._init();

        // 判断目标区域是否在当前可视区域内
        const { rowsIndex, colsIndex } = targetSheet.vi;
        const inView = rowsIndex.includes(rangeObj.s.r) && rowsIndex.includes(rangeObj.e.r)
                    && colsIndex.includes(rangeObj.s.c) && colsIndex.includes(rangeObj.e.c);

        // 不在可视区域内才滚动
        if (!inView) {
            targetSheet.vi.rowArr[0] = Math.max(0, rangeObj.s.r - 2);
            targetSheet.vi.colArr[0] = Math.max(0, rangeObj.s.c - 1);
        }

        // 设置选区
        targetSheet.activeCell = { r: rangeObj.s.r, c: rangeObj.s.c };
        targetSheet.activeAreas = [rangeObj];


        if (targetSheet !== SN.activeSheet) {
            SN.activeSheet = targetSheet;  // setter 自动 r()
        } else {
            SN._r();
        }
    };

    const renderDropdown = () => {
        const names = Object.entries(SN._definedNames || {});
        if (!names.length) {
            areaDropdown.innerHTML = `<div class="sn-area-dropdown-empty">${SN.t('core.canvas.canvas.areaDropdown.empty')}</div>`;
            return;
        }
        areaDropdown.innerHTML = names.map(([name, ref]) => `
            <div class="sn-area-dropdown-item" data-name="${name}" data-ref="${ref}">
                <span class="sn-area-dropdown-name">${name}</span>
                <span class="sn-area-dropdown-ref">${ref}</span>
            </div>
        `).join('');
    };

    // 点击箭头展开下拉
    areaBox.querySelector('.sn-area-arrow')?.addEventListener('click', (e) => {
        e.stopPropagation();
        areaBox.classList.toggle('open');
        if (areaBox.classList.contains('open')) {
            renderDropdown();
        }
    });

    areaDropdown.addEventListener('click', (e) => {
        const item = e.target.closest('.sn-area-dropdown-item');
        if (!item) return;
        gotoRef(item.dataset.ref);
        closeDropdown();
    });


    areaInput.addEventListener('blur', () => {
        const val = areaInput.value.trim();
        if (!val) return;
        // 检查是否是定义名称
        const ref = SN._definedNames?.[val];
        if (ref) {
            gotoRef(ref);
        } else if (/^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(val) || val.includes('!')) {
            // 尝试作为区域解析
            gotoRef(val);
        }
    });

    areaInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            areaInput.blur();
        }
        if (e.key === 'Escape') {
            closeDropdown();
            areaInput.blur();
        }
    });

    areaInput.addEventListener('focus', () => {
        areaInput.select();
    });

    // 点击外部关闭
    document.addEventListener('click', (e) => {
        if (!areaBox.contains(e.target)) {
            closeDropdown();
        }
    });
};

/**
 * Rendering trace arrow
 * @param {Sheet} sheet - Sheet Objects
 * @returns {void}
 */
Canvas.prototype.renderTraceArrows = function(sheet) {
    const arrows = sheet._traceArrows;
    if (!arrows || (arrows.precedents.length === 0 && arrows.dependents.length === 0)) return;

    const ctx = this.ctx;
    const scale = (this.dpr || 1) * (sheet.zoom ?? 1);

    ctx.save();


    ctx.strokeStyle = '#4472C4';
    ctx.fillStyle = '#4472C4';
    ctx.lineWidth = 1.5 * scale;
    arrows.precedents.forEach(arrow => {
        this.drawTraceArrow(ctx, sheet, arrow.from, arrow.to, scale);
    });

    // 绘制从属箭头（红色）
    ctx.strokeStyle = '#C44444';
    ctx.fillStyle = '#C44444';
    ctx.lineWidth = 1.5 * scale;
    arrows.dependents.forEach(arrow => {
        this.drawTraceArrow(ctx, sheet, arrow.from, arrow.to, scale);
    });

    ctx.restore();
};

/**
 * Draw a single trace arrow
 * @param {CanvasRenderingContext2D} ctx - Drawing Context
 * @param {Sheet} sheet - Sheet Objects
 * @ param {{r: number, c: number}} from - starting cell
 * @ param {{r: number, c: number}} to - end cell
 * @param {number} scale - Scaling coefficient
 * @returns {void}
 */
Canvas.prototype.drawTraceArrow = function(ctx, sheet, from, to, scale) {
    const fromInfo = sheet.getCellInViewInfo(from.r, from.c, false);
    const toInfo = sheet.getCellInViewInfo(to.r, to.c, false);

    // 至少有一端在视图内才绘制
    if (!fromInfo.inView && !toInfo.inView) return;


    const x1 = (fromInfo.x + fromInfo.w / 2) * scale;
    const y1 = (fromInfo.y + fromInfo.h / 2) * scale;
    const x2 = (toInfo.x + toInfo.w / 2) * scale;
    const y2 = (toInfo.y + toInfo.h / 2) * scale;

    // 绘制线条
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();

    // 绘制起点圆点
    ctx.beginPath();
    ctx.arc(x1, y1, 4 * scale, 0, Math.PI * 2);
    ctx.fill();

    // 绘制箭头头部
    const angle = Math.atan2(y2 - y1, x2 - x1);
    const headLen = 10 * scale;
    ctx.beginPath();
    ctx.moveTo(x2, y2);
    ctx.lineTo(x2 - headLen * Math.cos(angle - Math.PI / 6), y2 - headLen * Math.sin(angle - Math.PI / 6));
    ctx.lineTo(x2 - headLen * Math.cos(angle + Math.PI / 6), y2 - headLen * Math.sin(angle + Math.PI / 6));
    ctx.closePath();
    ctx.fill();
};
