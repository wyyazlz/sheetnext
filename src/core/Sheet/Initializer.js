import Row from '../Row/Row.js'
import Col from '../Col/Col.js'
import Drawing from '../Drawing/Drawing.js'
import Sparkline from '../Sparkline/Sparkline.js'
import CF from '../CF/CF.js'
import { parseSheetTables } from '../Table/xmlBuilder.js'
import { parseSheetSlicers } from '../Slicer/xmlParser.js'
import { Slicer } from '../Slicer/index.js'
import { PivotTable } from '../PivotTable/index.js'
import { _formatDate, _formatTime } from './helpers.js'
import { parseSheetProtectionNode } from './SheetProtection.js'

function _getActiveSelection(sheetView) {
    const selection = sheetView?.selection;
    const list = Array.isArray(selection) ? selection : (selection ? [selection] : []);
    const activePane = sheetView?.pane?.['_$activePane'];
    return list.find(item => item?._$pane === activePane) || list.find(item => !item?._$pane) || list[0] || {};
}

function _resolveSheetViewState(sheetView, Utils) {
    const pane = sheetView?.pane || {};
    const paneState = pane['_$state'] || '';
    const isFrozen = paneState === 'frozen' || paneState === 'frozenSplit';
    const splitCols = Math.max(0, Number(pane['_$xSplit']) || 0);
    const splitRows = Math.max(0, Number(pane['_$ySplit']) || 0);
    const sheetTopLeft = Utils.cellStrToNum(sheetView?.['_$topLeftCell'] ?? 'A1');
    const paneFallback = {
        r: splitRows > 0 ? sheetTopLeft.r + splitRows : sheetTopLeft.r,
        c: splitCols > 0 ? sheetTopLeft.c + splitCols : sheetTopLeft.c
    };
    const paneTopLeft = Utils.cellStrToNum(pane['_$topLeftCell'] ?? Utils.cellNumToStr(paneFallback));
    const freezeStartRow = isFrozen && splitRows > 0 ? sheetTopLeft.r : 0;
    const freezeStartCol = isFrozen && splitCols > 0 ? sheetTopLeft.c : 0;
    return {
        viewStart: isFrozen ? paneTopLeft : sheetTopLeft,
        freezeStartRow,
        freezeStartCol,
        frozenRows: isFrozen ? freezeStartRow + splitRows : 0,
        frozenCols: isFrozen ? freezeStartCol + splitCols : 0
    };
}

/**
 * Initialization module
 * Responsible for the initialization of sheets
 */

/**
 * Initialization module
 * Responsible for the initialization of sheets
 */

/**
 * Initialise Basic Configuration
 */
function initBasicConfig() {
    // 获取sheetXmlObj
    this._xmlObj = this.SN.Xml._xGetSheetXmlObj(this.rId);
    // 网格线
    this.showGridLines = this._xmlObj?.sheetViews?.sheetView?.[0]?.['_$showGridLines'] == '0' ? false : true
    // 行号列标
    this.showRowColHeaders = this._xmlObj?.sheetViews?.sheetView?.[0]?.['_$showRowColHeaders'] == '0' ? false : true
    // 缩放比例（Excel: 10%-400%）
    const view = this._xmlObj?.sheetViews?.sheetView?.[0] || {};
    const zoomScale = Number(view['_$zoomScale'] ?? view['_$zoomScaleNormal'] ?? 100);
    this._zoom = Math.min(4, Math.max(0.1, (Number.isFinite(zoomScale) ? zoomScale : 100) / 100));
    // 行高列宽
    this.defaultColWidth = Math.round((Number(this._xmlObj?.sheetFormatPr['_$defaultColWidth']) || 9.5) * 7.5);
    this.defaultRowHeight = Math.round((Number(this._xmlObj?.sheetFormatPr['_$defaultRowHeight']) || 15) * 1.333);
    // 视图配置
    this.views = this._xmlObj?.sheetViews?.sheetView ?? [{ pane: {} }];
    const viewState = _resolveSheetViewState(this.views[0], this.SN.Utils);
    this._frozenCols = viewState.frozenCols;
    this._frozenRows = viewState.frozenRows;
    this._freezeStartRow = viewState.freezeStartRow;
    this._freezeStartCol = viewState.freezeStartCol;
    const outlinePrNode = this._xmlObj?.sheetPr?.outlinePr;
    this.outlinePr = {
        applyStyles: outlinePrNode?._$applyStyles === '1',
        summaryBelow: outlinePrNode?._$summaryBelow !== '0',
        summaryRight: outlinePrNode?._$summaryRight !== '0',
        showOutlineSymbols: outlinePrNode?._$showOutlineSymbols !== '0'
    };
    // 打印设置
    if (this.SN.Print?.initSheetPrintSettings) {
        this.SN.Print.initSheetPrintSettings(this);
    }
}

/**
 * Initialise Sheet Protection
 */
function initSheetProtection() {
    const protection = parseSheetProtectionNode(this._xmlObj?.sheetProtection);
    if (protection) {
        this.protection.enable(protection);
    }
}

/**
 * Initialised Column Data
 */
function initRowsCols() {
    // 获取行列数
    const dimension = this._xmlObj?.dimension._$ref ?? "";
    let cc = 32; // 默认列数
    let rc = 200; // 默认行数
    if (dimension.includes(":")) {
        const range = this.rangeStrToNum(dimension);
        cc = range.e.c + 1
        rc = range.e.r + 1
    }

    // 初始化已有列
    (this._xmlObj?.cols?.col ?? []).forEach(col => {
        for (let i = parseInt(col['_$min']); i <= parseInt(col['_$max']); i++) { // 循环处理每个列的最小和最大值范围
            this.cols[i - 1] = new Col(col, this, i - 1);
        }
    });

    // 初始化行
    const sheetFile = this.SN.Xml._getSheetFileName(this.rId);
    const sheetKey = sheetFile ? `xl/worksheets/${sheetFile}` : null;
    const pendingRows = sheetKey ? this.SN.Xml._sheetDataRowsByPath?.[sheetKey] : null;
    const pendingMeta = sheetKey ? this.SN.Xml._sheetDataRowMetaByPath?.[sheetKey] : null;
    if (pendingRows) {
        this._setPendingXlsxRows(pendingRows, pendingMeta);
    } else {
        const rows = this._xmlObj?.sheetData?.row ?? [];
        (Array.isArray(rows) ? rows : [rows]).forEach(row => {
            const rIndex = parseInt(row['_$r']) - 1;
            if (rIndex == -1) console.log(rIndex)
            this.rows[rIndex] = new Row(row, this, rIndex);
        });
    }

    const sheetView = this.views[0] || {};
    const selection = _getActiveSelection(sheetView);
    const viewState = _resolveSheetViewState(sheetView, this.SN.Utils);
    this.activeCell = this.SN.Utils.cellStrToNum(selection?.['_$activeCell'] ?? 'A1'); // 活动的单元格
    this.vi = { colArr: [viewState.viewStart.c], rowArr: [viewState.viewStart.r] } // 起始行列

    if (this.activeCell.r > rc - 1) rc = this.activeCell.r + 1;
    if (this.activeCell.c > cc - 1) cc = this.activeCell.c + 1;
    this._virtualColCount = Math.max(this._virtualColCount || 0, cc);
    this._virtualRowCount = Math.max(this._virtualRowCount || 0, rc);
}

/**
 * Initialize Hyperlink
 */
function initHyperlinks() {
    const hyperlinks = this._xmlObj?.hyperlinks?.hyperlink ?? []
    const targetInfo = this.SN.Xml.resolveHyperlinkById(this.rId);
    hyperlinks.forEach(link => {
        const cell = this.getCell(link._$ref); // 获取单元格实例
        // 更新单元格的超链接信息
        cell.hyperlink = {
            target: targetInfo?.[link['_$r:id']]?._$Target || null, // 解析 r:id 对应的 URL
            location: link['_$location'] || null, // 内部位置跳转
            tooltip: link['_$tooltip'] || null // 鼠标悬停提示
        };
        if (link.display) cell.editVal = link.display;
    });
}

/**
 * Initializing data validation
 */
function initDataValidation() {
    const dataValidation = this._xmlObj?.dataValidations?.dataValidation ?? []
    dataValidation.forEach(val => {
        if (!val._$sqref.includes(':')) val._$sqref = `${val._$sqref}:${val._$sqref}`
        this.eachCells(val._$sqref, (r, c) => {
            const d = {
                type: val._$type, // 类型 list（一组离散有效值）whole（必须是整数）decimal（必须是十进制数）textLength（长度受控）custom（自定义公式控制）
                imeMode: val._$imeMode, // 输入法
                operator: val._$operator, // 运算符
                allowBlank: val._$allowBlank == '1' ? true : false,
                formula1: val.formula1,
                showDropDown: val._$showDropDown == '1' ? false : true, // 显示下拉按钮，仅序列生效，注意这里excel有误导
                formula2: val.formula2,
                showInputMessage: val._$showInputMessage == '1' ? true : false,
                promptTitle: val._$promptTitle, // 添加提示框标题
                prompt: val._$prompt,
                showErrorMessage: val._$showErrorMessage == '1' ? true : false,
                errorTitle: val._$errorTitle, // 添加错误框标题
                error: val._$error,
                errorStyle: val._$errorStyle
            }
            if (d.type == 'list') d.formula1 = val?.formula1?.slice(1, -1).split(","); // 是序列必须加下拉内容所以无需判断

            if (d.type == 'date') {
                d.formula1 = _formatDate(d.formula1)
                d.formula2 = _formatDate(d.formula2)
            }
            if (d.type == 'time') {
                d.formula1 = _formatTime(d.formula1)
                d.formula2 = _formatTime(d.formula2)
            }

            this.getCell(r, c)._dataValidation = d;
        })
    });
}

/**
 * Initialise Merge Cells
 */
function initMergeCells() {
    (this._xmlObj?.mergeCells?.mergeCell ?? []).forEach(item => {
        const rg = this.rangeStrToNum(item['_$ref'])
        this.merges.push(rg);
        this.eachCells(rg, (r, c) => {
            const cell = this.getCell(r, c);
            cell.isMerged = true;
            cell.master = { r: rg.s.r, c: rg.s.c }
        })
    });
}

/**
 * Initializing Charts and Graphics
 */
function initDrawings() {
    // 创建图纸管理器（复刻 Sparkline 架构）
    this.Drawing = new Drawing(this);
    this.Drawing._parse(this._xmlObj);
}

/**
 * Initialise Slicer
 */
function initSlicers() {
    this.Slicer = new Slicer(this);
    parseSheetSlicers(this, this.SN.Xml);
}

/**
 * Initializing Periscope Manager
 */
function initPivotTables() {
    if (!this.PivotTable) {
        this.PivotTable = new PivotTable(this);
        return;
    }

    if (this.PivotTable.size === 0) return;

    this.PivotTable.getAll().forEach(pt => {
        const isEmptyPivot = pt.rowFields.length === 0 &&
            pt.colFields.length === 0 &&
            pt.dataFields.length === 0 &&
            pt.pageFields.length === 0;
        if (isEmptyPivot) return;
        pt._needsRecalculate = true;
        pt.calculate();
        pt.render({ silent: true });
    });
}

/**
 * Initializing mini-maps
 */
function initSparklines() {
    this.Sparkline = new Sparkline(this);
    this.Sparkline.parse(this._xmlObj);
}

/**
 * Initialise Notes
 */
function initComments() {
    this.Comment.parse(this._xmlObj);
}

/**
 * Initialise Conditional Formatting
 */
function initCF() {
    this.CF = new CF(this);
    this.CF.parse(this._xmlObj);
}

/**
 * Initialize AutoFilter
 */
function initAutoFilter() {
    this.AutoFilter.parse(this._xmlObj);
}

/**
 * Initializing Super Table
 */
function initTables() {
    // 获取工作表文件名
    const sheetFile = this.SN.Xml._getSheetFileName(this.rId);
    if (sheetFile) {
        parseSheetTables(this, this.SN.Xml.obj, sheetFile);
    }

}

/**
 * Initialization of the dependency chain
 */
function initDependencyGraph() {
    // 【依赖链集成】初始化完成后，建立所有公式的依赖关系
    if (this.SN.DependencyGraph) {
        this.SN.DependencyGraph.rebuildForSheet(this);
    }
}

/**
 * Main Initialization Method
 */
export function _init() {
    if (this.initialized) return this
    this.initialized = true

    // 依次执行各个初始化步骤
    initBasicConfig.call(this);
    initSheetProtection.call(this);
    initRowsCols.call(this);
    initHyperlinks.call(this);
    initDataValidation.call(this);
    initMergeCells.call(this);
    initDrawings.call(this);
    initSparklines.call(this);
    initComments.call(this);
    initCF.call(this);
    initAutoFilter.call(this);
    initTables.call(this);
    initPivotTables.call(this);
    initSlicers.call(this);
    initDependencyGraph.call(this);
    if (this.SN.calcMode !== 'manual') {
        this.SN.DependencyGraph?.recalculateSheetFormulas?.(this);
    }

    return this
}
