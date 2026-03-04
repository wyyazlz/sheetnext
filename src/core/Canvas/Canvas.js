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
 * 画布渲染与交互管理
 * @title 🖱️ 绑定操作
 * @class
 */
export default class Canvas {

    /**
     * @param {Object} SN - SheetNext 主实例
     * @param {Object} license
     */
    constructor(SN, license) {
        /**
         * SheetNext 主实例
         * @type {Object}
         */
        this.SN = SN
        _clm.set(this, license);
        /**
         * 工具方法集合
         * @type {Utils}
         */
        this.Utils = SN.Utils
        /**
         * 画布容器
         * @type {HTMLElement}
         */
        this.dom = SN.containerDom.querySelector('.sn-canvas')

        /**
         * 显示层画布
         * @type {HTMLCanvasElement}
         */
        this.showLayer = SN.containerDom.querySelector(".sn-show-layer")

        /**
         * 操作层画布
         * @type {HTMLCanvasElement}
         */
        this.handleLayer = SN.containerDom.querySelector(".sn-handle-layer")
        // 缓冲画布
        /**
         * 缓冲画布
         * @type {HTMLCanvasElement}
         */
        this.buffer = document.createElement('canvas');
        /**
         * 快照画布
         * @type {HTMLCanvasElement}
         */
        this.showLayerKZ = document.createElement('canvas');
        /**
         * 画笔预览画布
         * @type {HTMLCanvasElement}
         */
        this.pCanvas = document.createElement('canvas');
        this.pCanvas.width = this.pCanvas.height = 8;
        /**
         * 画笔预览上下文
         * @type {CanvasRenderingContext2D}
         */
        this.pCtx = this.pCanvas.getContext('2d');
        // 其他画布信息
        /**
         * 当前激活边框信息
         * @type {Object|null}
         */
        this.activeBorderInfo = null;
        /**
         * 多选边框信息
         * @type {Object}
         */
        this.mpBorder = {}
        /**
         * 最近一次 padding 信息
         * @type {Object|null}
         */
        this.lastPadding = null;
        // 鼠标事件
        /**
         * 鼠标指针当前指向区域
         * @type {string|null}
         */
        this.moveAreaName = null; // 鼠标指针当前指向区域 row-col-resize drawing fill-handle active-border
        /**
         * 是否禁用移动检测
         * @type {boolean}
         */
        this.disableMoveCheck = false;
        /**
         * 是否禁用滚动
         * @type {boolean}
         */
        this.disableScroll = false; // 是否禁用滚动功能
        // 必要DOM（初始化时获取）
        /**
         * 输入编辑器
         * @type {HTMLElement}
         */
        this.input = SN.containerDom.querySelector('.sn-input-editor'); // 画布富文本输入框
        /**
         * 公式栏输入框
         * @type {HTMLElement}
         */
        this.formulaBar = SN.containerDom.querySelector('.sn-formula-bar'); //工具栏输入框
        this.inputEditing = false;
        /**
         * 区域输入框
         * @type {HTMLInputElement}
         */
        this.areaInput = SN.containerDom.querySelector('.sn-area-input');
        /**
         * 区域选择容器
         * @type {HTMLElement}
         */
        this.areaBox = SN.containerDom.querySelector('.sn-area-box');
        /**
         * 区域下拉容器
         * @type {HTMLElement}
         */
        this.areaDropdown = SN.containerDom.querySelector('.sn-area-dropdown');
        /**
         * 横向滚动条
         * @type {HTMLElement}
         */
        this.rollX = SN.containerDom.querySelector('.sn-roll-x')
        /**
         * 横向滚动条滑块
         * @type {HTMLElement}
         */
        this.rollXS = SN.containerDom.querySelector('.sn-roll-x>div')
        /**
         * 纵向滚动条
         * @type {HTMLElement}
         */
        this.rollY = SN.containerDom.querySelector('.sn-roll-y')
        /**
         * 纵向滚动条滑块
         * @type {HTMLElement}
         */
        this.rollYS = SN.containerDom.querySelector('.sn-roll-y>div')
        /**
         * 图形容器
         * @type {HTMLElement}
         */
        this.drawingsCon = SN.containerDom.querySelector('.sn-drawings-con')
        /**
         * 切片器容器
         * @type {HTMLElement}
         */
        this.slicersCon = SN.containerDom.querySelector('.sn-slicers-con')

        /**
         * 最大纵向滚动距离
         * @type {number}
         */
        this.maxTop = 0;
        /**
         * 最大横向滚动距离
         * @type {number}
         */
        this.maxLeft = 0;

        /**
         * 懒加载 DOM 管理器
         * @type {LazyDOM}
         */
        this.lazyDOM = new LazyDOM(SN);

        /**
         * 筛选面板管理器
         * @type {FilterPanel}
         */
        this.filterPanelManager = new FilterPanel(this);

        /**
         * 渲染计数
         * @type {number}
         */
        this.rCount = 0

        /**
         * 当前悬停批注
         * @type {Object|null}
         */
        this.hoveredComment = null;

        /**
         * 统计求和：DOM
         * @type {HTMLElement|null}
         */
        this.statSum = SN.containerDom.querySelector('[data-stat="sum"]');
        /**
         * 统计计数：DOM
         * @type {HTMLElement|null}
         */
        this.statCount = SN.containerDom.querySelector('[data-stat="count"]');
        /**
         * 统计平均：DOM
         * @type {HTMLElement|null}
         */
        this.statAvg = SN.containerDom.querySelector('[data-stat="avg"]');
        this._statRafId = null;

        /**
         * 高亮颜色，null 表示关闭
         * @type {string|null}
         */
        this.highlightColor = null; // 高亮颜色，null 表示关闭
        // 护眼模式
        /**
         * 是否开启护眼模式
         * @type {boolean}
         */
        this.eyeProtectionMode = false;
        /**
         * 护眼颜色
         * @type {string}
         */
        this.eyeProtectionColor = 'rgb(204,232,207)';
        //事件监听
        this.addEventListeners();
        this._initAreaSelector();
        this.viewResize()
    }


    /**
     * 内边距类型
     * @type {string}     */
    get paddingType() { return this.lazyDOM.paddingType; }
    /**
     * 超链接跳转按钮
     * @type {HTMLElement|null}     */
    get hyperlinkJumpBtn() { return this.lazyDOM.hyperlinkBtn; }
    /**
     * 超链接提示
     * @type {HTMLElement|null}     */
    get hyperlinkTip() { return this.lazyDOM.hyperlinkTip; }
    /**
     * 数据验证选择器
     * @type {HTMLElement|null}     */
    get validSelect() { return this.lazyDOM.validSelect; }
    /**
     * 数据验证提示
     * @type {HTMLElement|null}     */
    get validTip() { return this.lazyDOM.validTip; }
    /**
     * 滚动提示
     * @type {HTMLElement|null}     */
    get scrollTip() { return this.lazyDOM.scrollTip; }

    /**
     * 更新画布缩放/长宽后重新获取 dom 长高
     * @returns {void}
     */
    viewResize() {
        this.dpr = window.devicePixelRatio || 1;
        const { clientWidth: w, clientHeight: h } = this.dom;
        [this.showLayer, this.handleLayer, this.buffer, this.showLayerKZ].forEach(canvas => {
            canvas.width = w * this.dpr;
            canvas.height = h * this.dpr;
            canvas.style.width = w + 'px'
            canvas.style.height = h + 'px'
        })

        this._width = null
        this._height = null

        this.ctx = this.showLayer.getContext('2d');
        this.ctxKZ = this.showLayerKZ.getContext('2d'); // 快照1
        this.HDctx = this.handleLayer.getContext('2d');

        this.#initBufferContext();

        // 窗口变化时更新滚动条尺寸
        if (this.activeSheet) {
            this.updateScrollBarSize();
            this.updateOverlayContainers();
        }
    }

    /**
     * 可视宽度
     * @type {number}     */
    get width() {
        return this._width ?? (this._width = this.showLayer.offsetWidth)
    }

    /**
     * 可视高度
     * @type {number}     */
    get height() {
        return this._height ?? (this._height = this.showLayer.offsetHeight)
    }

    /**
     * 逻辑视图宽度
     * @type {number}     */
    get viewWidth() {
        return this.width / (this.activeSheet?.zoom ?? 1);
    }

    /**
     * 逻辑视图高度
     * @type {number}     */
    get viewHeight() {
        return this.height / (this.activeSheet?.zoom ?? 1);
    }

    /**
     * 物理像素转逻辑坐标
     * @param {number} value - 物理像素
     * @returns {number}
     */
    toLogical(value) {
        return value / (this.activeSheet?.zoom ?? 1);
    }

    /**
     * 逻辑坐标转物理像素
     * @param {number} value - 逻辑坐标
     * @returns {number}
     */
    toPhysical(value) {
        return value * (this.activeSheet?.zoom ?? 1);
    }

    /**
     * 像素比
     * @type {number}     */
    get pixelRatio() {
        return this._pixelRatio ?? ((this.dpr || 1) * (this.activeSheet?.zoom ?? 1));
    }

    /**
     * 半像素偏移
     * @type {number}     */
    get halfPixel() {
        return 0.5 / this.pixelRatio;
    }

    /**
     * 转换为像素比下的值
     * @param {number} value - 原始值
     * @returns {number}
     */
    px(value) {
        return value / this.pixelRatio;
    }

    /**
     * 对齐像素比
     * @param {number} value - 原始值
     * @returns {number}
     */
    snap(value) {
        const ratio = this.pixelRatio;
        return Math.round(value * ratio) / ratio;
    }

    /**
     * 对齐矩形到像素比
     * @param {number} x - X
     * @param {number} y - Y
     * @param {number} w - 宽
     * @param {number} h - 高
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
     * 当前活动工作表
     * @type {Sheet}     */
    get activeSheet() {
        return this.SN.activeSheet
    }

    /**
     * 是否处于输入编辑状态
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
     * 获取事件在画布中的位置（逻辑坐标）
     * @param {MouseEvent|{clientX:number, clientY:number}} event - 鼠标事件或带坐标对象
     * @returns {{x: number, y: number}}
     */
    getEventPosition(event) {
        const rect = this.handleLayer.getBoundingClientRect();
        const x = this.toLogical(event.clientX - rect.left);
        const y = this.toLogical(event.clientY - rect.top);
        return { x, y };
    }

    /**
     * 通过位置获取单元格索引
     * @param {number} x - 逻辑坐标 X
     * @param {number} y - 逻辑坐标 Y
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
     * 应用缩放比例并同步滚动条与输入框位置
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
     * 清除缓冲画布（物理像素）
     * @param {number} [width=this.buffer.width] - 宽度
     * @param {number} [height=this.buffer.height] - 高度
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
     * 更新滚动条尺寸
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
     * 更新覆盖层容器尺寸
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
     * 更新滚动条位置
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
     * 获取事件对应的单元格索引
     * @param {MouseEvent|{clientX:number, clientY:number}} event - 鼠标事件或带坐标对象
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
     * 设置高亮颜色并同步工具栏状态
     * @param {MouseEvent} event - 触发事件
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
     * 切换护眼模式
     * @returns {void}
     */
    toggleEyeProtection() {
        this.eyeProtectionMode = !this.eyeProtectionMode;


        const btn = this.SN.containerDom.querySelector('[data-icon="huyan_"]')?.closest('.sn-tools-item');
        if (btn) btn.classList.toggle('sn-btn-selected', this.eyeProtectionMode);

        this.r();
    }



    /**
     * 渲染画布（支持截图模式）
     * @param {string} [type] - 渲染类型，'s' 为快照渲染
     * @param {Object} [options] - 渲染选项
     * @param {HTMLCanvasElement} [options.targetCanvas] - 目标画布（截图模式）
     * @param {number} [options.width] - 截图宽度
     * @param {number} [options.height] - 截图高度
     * @param {number} [options.dpr] - 指定 dpr
     * @param {Sheet} [options.sheet] - 指定工作表
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
 * 渲染追踪箭头
 * @param {Sheet} sheet - 工作表对象
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
 * 绘制单个追踪箭头
 * @param {CanvasRenderingContext2D} ctx - 绘图上下文
 * @param {Sheet} sheet - 工作表对象
 * @param {{r:number, c:number}} from - 起点单元格
 * @param {{r:number, c:number}} to - 终点单元格
 * @param {number} scale - 缩放系数
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
