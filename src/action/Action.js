// 对系统工具栏等按钮封装对应方法

// 导入模块
import * as RowColActions from './RowCol.js';
import * as FormatActions from './Format.js';
import * as StyleActions from './Style.js';
import * as InsertActions from './Insert.js';
import * as WorkbookActions from './Workbook.js';
import * as StorageActions from './Storage.js';
import * as ClipboardActions from './Clipboard.js';
import * as CondFormatActions from './CondFormat.js';
import * as FormulaActions from './Formula.js';
import * as NameManagerActions from './NameManager.js';
import * as FormulaAuditActions from './FormulaAudit.js';
import * as FilterActions from './Filter.js';
import * as FindActions from './Find.js';
import * as DataActions from './Data.js';
import * as PivotTableActions from './PivotTable.js';
import * as TableActions from './Table.js';
import * as ViewActions from './View.js';
import * as PrintActions from './Print.js';
import * as ProtectionActions from './Protection.js';

export default class Fun {

    constructor(SN) {
        this.SN = SN
    }

}

// 使用原型链挂载行列操作模块方法
Object.assign(Fun.prototype, {
    // 右键菜单兼容
    hidRows: RowColActions.hidRows,
    delRows: RowColActions.delRows,
    addRows: RowColActions.addRows,
    delCols: RowColActions.delCols,
    hidCols: RowColActions.hidCols,
    addCols: RowColActions.addCols,
    // 工具栏方法
    insertRowDialog: RowColActions.insertRowDialog,
    insertColDialog: RowColActions.insertColDialog,
    deleteRowAction: RowColActions.deleteRowAction,
    deleteColAction: RowColActions.deleteColAction,
    hideRowAction: RowColActions.hideRowAction,
    hideColAction: RowColActions.hideColAction,
    unhideRowAction: RowColActions.unhideRowAction,
    unhideColAction: RowColActions.unhideColAction,
    unhideAllRowsAction: RowColActions.unhideAllRowsAction,
    unhideAllColsAction: RowColActions.unhideAllColsAction,
    applyRowColInsertByPos: RowColActions.applyRowColInsertByPos,
    setRowHeightDialog: RowColActions.setRowHeightDialog,
    setColWidthDialog: RowColActions.setColWidthDialog,
    setDefaultRowHeightDialog: RowColActions.setDefaultRowHeightDialog,
    setDefaultColWidthDialog: RowColActions.setDefaultColWidthDialog,
    autoFitRowHeight: RowColActions.autoFitRowHeight,
    autoFitColWidth: RowColActions.autoFitColWidth,
    groupRows: RowColActions.groupRows,
    ungroupRows: RowColActions.ungroupRows
});

// 使用原型链挂载格式操作模块方法
Object.assign(Fun.prototype, {
    decimalPlaces: FormatActions.decimalPlaces,
    areasNumFmt: FormatActions.areasNumFmt,
    customNumFmtDialog: FormatActions.customNumFmtDialog,
    clearAreaFormat: FormatActions.clearAreaFormat,
    textToNumber: FormatActions.textToNumber,
    numberToText: FormatActions.numberToText,
    toLowercase: FormatActions.toLowercase,
    toUppercase: FormatActions.toUppercase,
    trimSpaces: FormatActions.trimSpaces,
    textToColumns: FormatActions.textToColumns,
    flashFill: FormatActions.flashFill
});

// 使用原型链挂载样式操作模块方法
Object.assign(Fun.prototype, {
    fillColorBtn: StyleActions.fillColorBtn,
    fillColorChange: StyleActions.fillColorChange,
    borderColorBtn: StyleActions.borderColorBtn,
    borderColorChange: StyleActions.borderColorChange,
    fontColorBtn: StyleActions.fontColorBtn,
    fontColorChange: StyleActions.fontColorChange,
    applyFontColor: StyleActions.applyFontColor,
    applyFillColor: StyleActions.applyFillColor,
    applyBorderPreset: StyleActions.applyBorderPreset,
    applyBorderStyle: StyleActions.applyBorderStyle,
    initToolbarDefaults: StyleActions.initToolbarDefaults,
    areasStyle: StyleActions.areasStyle,
    fontInversion: StyleActions.fontInversion,
    toggleWrapText: StyleActions.toggleWrapText,
    changeFontSize: StyleActions.changeFontSize,
    changeIndent: StyleActions.changeIndent,
    setTextRotation: StyleActions.setTextRotation,
    applyCellStyle: StyleActions.applyCellStyle
});

// 使用原型链挂载插入操作模块方法
Object.assign(Fun.prototype, {
    insertImages: InsertActions.insertImages,
    insertAreaScreenshot: InsertActions.insertAreaScreenshot,
    insertChart: InsertActions.insertChart,
    openChartModal: InsertActions.openChartModal,
    insertComment: InsertActions.insertComment,
    insertHyperlink: InsertActions.insertHyperlink,
    insertTextBox: InsertActions.insertTextBox,
    insertSparkline: InsertActions.insertSparkline,
    openSparklineDialog: InsertActions.openSparklineDialog
});

// 使用原型链挂载工作簿操作模块方法
Object.assign(Fun.prototype, {
    fullScreen: WorkbookActions.fullScreen,
    moreSheet: WorkbookActions.moreSheet,
    rename: WorkbookActions.rename,
    newWorkbook: WorkbookActions.newWorkbook,
    confirmNewWorkbook: WorkbookActions.confirmNewWorkbook,
    importFromUrl: WorkbookActions.importFromUrl,
    showProperties: WorkbookActions.showProperties
});

// 使用原型链挂载视图操作模块方法
Object.assign(Fun.prototype, {
    zoomCustom: ViewActions.zoomCustom,
    // 冻结窗格
    freezeRows: ViewActions.freezeRows,
    freezeCols: ViewActions.freezeCols,
    freezeRowsAndCols: ViewActions.freezeRowsAndCols,
    freezeAtActiveCell: ViewActions.freezeAtActiveCell,
    freezeToCurrentRow: ViewActions.freezeToCurrentRow,
    freezeToCurrentCol: ViewActions.freezeToCurrentCol,
    unfreeze: ViewActions.unfreeze
});

// 使用原型链挂载本地存储模块方法
Object.assign(Fun.prototype, {
    saveToLocal: StorageActions.saveToLocal,
    loadFromLocal: StorageActions.loadFromLocal,
    hasLocalData: StorageActions.hasLocalData,
    clearLocalData: StorageActions.clearLocalData
});

// 使用原型链挂载剪贴板模块方法
Object.assign(Fun.prototype, {
    copy: ClipboardActions.copy,
    cut: ClipboardActions.cut,
    paste: ClipboardActions.paste,
    clearClipboard: ClipboardActions.clearClipboard,
    handlePasteEvent: ClipboardActions.handlePasteEvent
});

// 使用原型链挂载条件格式模块方法
Object.assign(Fun.prototype, {
    condFormat: CondFormatActions.condFormat
});

Object.assign(Fun.prototype, {
    openFormulaDialog: FormulaActions.openFormulaDialog,
    insertFunction: FormulaActions.insertFunction
});

Object.assign(Fun.prototype, {
    nameManager: NameManagerActions.nameManager,
    defineName: NameManagerActions.defineName,
    useInFormula: NameManagerActions.useInFormula
});

// 公式审核模块方法
Object.assign(Fun.prototype, {
    tracePrecedents: FormulaAuditActions.tracePrecedents,
    traceDependents: FormulaAuditActions.traceDependents,
    removeArrows: FormulaAuditActions.removeArrows,
    showFormulas: FormulaAuditActions.showFormulas,
    errorCheck: FormulaAuditActions.errorCheck,
    evaluateFormula: FormulaAuditActions.evaluateFormula,
    calculateNow: FormulaAuditActions.calculateNow,
    calculateSheet: FormulaAuditActions.calculateSheet,
    setCalcMode: FormulaAuditActions.setCalcMode
});

// 筛选模块方法
Object.assign(Fun.prototype, {
    toggleAutoFilter: FilterActions.toggleAutoFilter,
    clearFilter: FilterActions.clearFilter,
    reapplyFilter: FilterActions.reapplyFilter,
    sortAsc: FilterActions.sortAsc,
    sortDesc: FilterActions.sortDesc,
    customSort: FilterActions.customSort
});

// 数据工具模块方法
Object.assign(Fun.prototype, {
    subtotal: DataActions.subtotal,
    consolidate: DataActions.consolidate,
    dataValidationDialog: DataActions.dataValidationDialog,
    clearDataValidation: DataActions.clearDataValidation
});

// 查找模块方法
Object.assign(Fun.prototype, {
    openFindDialog: FindActions.openFindDialog
});

// 打印模块方法
Object.assign(Fun.prototype, {
    printPreview: PrintActions.printPreview,
    exportPdf: PrintActions.exportPdf,
    pageBreakPreview: PrintActions.pageBreakPreview,
    toggleShowPageBreaks: PrintActions.toggleShowPageBreaks,
    insertPageBreak: PrintActions.insertPageBreak,
    togglePrintGridlines: PrintActions.togglePrintGridlines,
    togglePrintHeadings: PrintActions.togglePrintHeadings,
    setPageMarginsPreset: PrintActions.setPageMarginsPreset,
    setPageMarginsDialog: PrintActions.setPageMarginsDialog,
    setPageOrientation: PrintActions.setPageOrientation,
    setPaperSize: PrintActions.setPaperSize,
    setPrintScale: PrintActions.setPrintScale,
    setPrintScaleDialog: PrintActions.setPrintScaleDialog,
    setPrintAreaFromSelection: PrintActions.setPrintAreaFromSelection,
    clearPrintArea: PrintActions.clearPrintArea,
    setPrintTitles: PrintActions.setPrintTitles,
    setPageBackground: PrintActions.setPageBackground,
    setHeaderFooter: PrintActions.setHeaderFooter
});

// 透视表模块方法
Object.assign(Fun.prototype, {
    createPivotTable: PivotTableActions.createPivotTable,
    refreshPivotTable: PivotTableActions.refreshPivotTable,
    openPivotTableOptions: PivotTableActions.openPivotTableOptions,
    showPivotReportFilterPages: PivotTableActions.showPivotReportFilterPages,
    // 分析面板
    renamePivotTable: PivotTableActions.renamePivotTable,
    pivotFieldSettings: PivotTableActions.pivotFieldSettings,
    changePivotDataSource: PivotTableActions.changePivotDataSource,
    clearPivotTable: PivotTableActions.clearPivotTable,
    movePivotTable: PivotTableActions.movePivotTable,
    insertPivotChart: PivotTableActions.insertPivotChart,
    recommendPivotTable: PivotTableActions.recommendPivotTable,
    insertPivotSlicer: PivotTableActions.insertPivotSlicer,
    togglePivotFieldList: PivotTableActions.togglePivotFieldList,
    togglePivotDrillButtons: PivotTableActions.togglePivotDrillButtons,
    togglePivotFieldHeaders: PivotTableActions.togglePivotFieldHeaders,
    expandPivotField: PivotTableActions.expandPivotField,
    collapsePivotField: PivotTableActions.collapsePivotField,
    // 设计面板
    setPivotSubtotals: PivotTableActions.setPivotSubtotals,
    setPivotGrandTotals: PivotTableActions.setPivotGrandTotals,
    setPivotLayout: PivotTableActions.setPivotLayout,
    setPivotRepeatLabels: PivotTableActions.setPivotRepeatLabels,
    setPivotBlankRows: PivotTableActions.setPivotBlankRows,
    togglePivotOption: PivotTableActions.togglePivotOption
});

// 超级表模块方法
Object.assign(Fun.prototype, {
    createTable: TableActions.createTable,
    formatAsTable: TableActions.formatAsTable,
    convertTableToRange: TableActions.convertTableToRange,
    resizeTable: TableActions.resizeTable,
    changeTableStyle: TableActions.changeTableStyle,
    toggleTableOption: TableActions.toggleTableOption,
    renameTable: TableActions.renameTable,
    deleteTable: TableActions.deleteTable,
    clearTableStyle: TableActions.clearTableStyle,
    insertTableTotalsFunction: TableActions.insertTableTotalsFunction,
    removeDuplicates: TableActions.removeDuplicates,
    insertSlicer: TableActions.insertSlicer
});

// 工作表保护模块方法
Object.assign(Fun.prototype, {
    protectSheet: ProtectionActions.protectSheet,
    unprotectSheet: ProtectionActions.unprotectSheet,
    configProtection: ProtectionActions.configProtection,
    cancelProtection: ProtectionActions.cancelProtection,
    configCellProtection: ProtectionActions.configCellProtection,
    setPermission: ProtectionActions.setPermission,
    resetPermissions: ProtectionActions.resetPermissions
});
