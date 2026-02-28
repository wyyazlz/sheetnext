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

/**
 * 初始化模块
 * 负责工作表的初始化工作
 */

/**
 * 初始化基础配置
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
    // 冻结信息
    if (this.views[0]?.pane?.['_$state'] == 'frozen') {
        this._frozenCols = Number(this.views[0].pane['_$xSplit']) || 0;
        this._frozenRows = Number(this.views[0].pane['_$ySplit']) || 0;
    }
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
 * 初始化工作表保护
 */
function initSheetProtection() {
    const protection = parseSheetProtectionNode(this._xmlObj?.sheetProtection);
    if (protection) {
        this.protection.enable(protection);
    }
}

/**
 * 初始化行列数据
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
    (this._xmlObj?.sheetData?.row ?? []).forEach(row => {
        const rIndex = parseInt(row['_$r']) - 1;
        if (rIndex == -1) console.log(rIndex)
        this.rows[rIndex] = new Row(row, this, rIndex);
    });

    this.activeCell = this.SN.Utils.cellStrToNum(this.views[0]?.selection?.['_$activeCell'] ?? 'A1'); // 活动的单元格
    const topLeftCell = this.SN.Utils.cellStrToNum(this.views[0]?.['_$topLeftCell'] ?? 'A1'); // 视图位置
    this.vi = { colArr: [topLeftCell.c], rowArr: [topLeftCell.r] } // 起始行列

    if (this.activeCell.r > rc - 1) rc = this.activeCell.r + 1;
    if (this.activeCell.c > cc - 1) cc = this.activeCell.c + 1;
    if (!this.cols[cc - 1]) this.cols[cc - 1] = new Col(null, this, cc - 1);
    if (!this.rows[rc - 1]) this.rows[rc - 1] = new Row(null, this, rc - 1);
}

/**
 * 初始化超链接
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
 * 初始化数据验证
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
 * 初始化合并单元格
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
 * 初始化图表和图形
 */
function initDrawings() {
    // 创建图纸管理器（复刻 Sparkline 架构）
    this.Drawing = new Drawing(this);
    this.Drawing._parse(this._xmlObj);
}

/**
 * 初始化切片器
 */
function initSlicers() {
    this.Slicer = new Slicer(this);
    parseSheetSlicers(this, this.SN.Xml);
}

/**
 * 初始化透视表管理器
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
 * 初始化迷你图
 */
function initSparklines() {
    this.Sparkline = new Sparkline(this);
    this.Sparkline.parse(this._xmlObj);
}

/**
 * 初始化批注
 */
function initComments() {
    this.Comment.parse(this._xmlObj);
}

/**
 * 初始化条件格式
 */
function initCF() {
    this.CF = new CF(this);
    this.CF.parse(this._xmlObj);
}

/**
 * 初始化自动筛选
 */
function initAutoFilter() {
    this.AutoFilter.parse(this._xmlObj);
}

/**
 * 初始化超级表
 */
function initTables() {
    // 获取工作表文件名
    const sheetFile = this.SN.Xml._getSheetFileName(this.rId);
    if (sheetFile) {
        parseSheetTables(this, this.SN.Xml.obj, sheetFile);
    }

}

/**
 * 初始化依赖链
 */
function initDependencyGraph() {
    // 【依赖链集成】初始化完成后，建立所有公式的依赖关系
    if (this.SN.DependencyGraph) {
        this.SN.DependencyGraph.rebuildForSheet(this);
    }
}

/**
 * 主初始化方法
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

    return this
}
