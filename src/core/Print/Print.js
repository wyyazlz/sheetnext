const PRINT_DPI = 96;

const PAPER_SIZES = {
    A4: { name: 'A4', paperSize: 9, widthIn: 8.27, heightIn: 11.69 },
    Letter: { name: 'Letter', paperSize: 1, widthIn: 8.5, heightIn: 11 }
};

const PAPER_SIZE_BY_ID = {
    9: 'A4',
    1: 'Letter'
};

const DEFAULT_MARGINS = {
    left: 0.7,
    right: 0.7,
    top: 0.75,
    bottom: 0.75,
    header: 0.3,
    footer: 0.3
};

function toNumber(val, fallback) {
    const num = Number(val);
    return Number.isFinite(num) ? num : fallback;
}

function toInt(val, fallback) {
    const num = Number.parseInt(val, 10);
    return Number.isFinite(num) ? num : fallback;
}

function toBool(val, fallback = false) {
    if (val === true || val === '1' || val === 1 || val === 'true') return true;
    if (val === false || val === '0' || val === 0 || val === 'false') return false;
    return fallback;
}

function normalizeArray(val) {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}

function clone(obj) {
    return JSON.parse(JSON.stringify(obj));
}

function clampScale(scale) {
    const num = toNumber(scale, 100);
    return Math.min(400, Math.max(10, Math.round(num)));
}

function escapeHtml(text) {
    return String(text ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function escapeAttr(text) {
    return escapeHtml(text)
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/**
 * Print Settings and Rendering
 * @title Print
 * @class
 */
export default class Print {
    /**
     * @param {Object} SN - SheetNext Main Instance
     */
    constructor(SN) {
        /**
         * SheetNext Main Instance
         * @type {Object}
         */
        this.SN = SN;
    }

    /**
     * Create Default Printing Settings
     * @returns {Object}
     */
    createDefaultSettings() {
        const paper = PAPER_SIZES.A4;
        return {
            pageSetup: {
                paperSize: paper.paperSize,
                paperName: paper.name,
                orientation: 'portrait',
                scale: 100,
                fitToWidth: 0,
                fitToHeight: 0
            },
            pageMargins: { ...DEFAULT_MARGINS },
            printOptions: {
                gridlines: false,
                headings: false
            },
            printArea: null,
            printTitles: {
                rows: null,
                cols: null
            },
            headerFooter: {
                header: '',
                footer: ''
            },
            pageBreaks: {
                rowBreaks: [],
                colBreaks: []
            },
            pageBackground: null
        };
    }

    /**
     * Get and ensure print settings for sheets
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    ensureSettings(sheet = this.SN.activeSheet) {
        if (!sheet) return null;
        if (!sheet.printSettings) {
            sheet.printSettings = this.createDefaultSettings();
        }
        return sheet.printSettings;
    }

    /**
     * Initialising Sheet Printing Settings
     * @param {Sheet} sheet - Sheet
     * @returns {Object}
     */
    initSheetPrintSettings(sheet) {
        const settings = this.readSettingsFromXml(sheet?._xmlObj, sheet);
        sheet.printSettings = settings;
        if (typeof sheet.showPageBreaks !== 'boolean') {
            sheet.showPageBreaks = false;
        }
        return settings;
    }

    /**
     * Read print settings from XML
     * @param {Object} xmlObj - Worksheet XML Object
     * @returns {Object}
     */
    readSettingsFromXml(xmlObj, sheet = null) {
        const settings = this.createDefaultSettings();
        if (!xmlObj) return settings;

        const pageMargins = xmlObj.pageMargins || {};
        settings.pageMargins = {
            left: toNumber(pageMargins._$left, settings.pageMargins.left),
            right: toNumber(pageMargins._$right, settings.pageMargins.right),
            top: toNumber(pageMargins._$top, settings.pageMargins.top),
            bottom: toNumber(pageMargins._$bottom, settings.pageMargins.bottom),
            header: toNumber(pageMargins._$header, settings.pageMargins.header),
            footer: toNumber(pageMargins._$footer, settings.pageMargins.footer)
        };

        const pageSetup = xmlObj.pageSetup || {};
        const paperSize = toInt(pageSetup._$paperSize, settings.pageSetup.paperSize);
        const paperName = PAPER_SIZE_BY_ID[paperSize] || settings.pageSetup.paperName;
        settings.pageSetup.paperSize = paperSize;
        settings.pageSetup.paperName = paperName;
        settings.pageSetup.orientation = pageSetup._$orientation === 'landscape' ? 'landscape' : 'portrait';

        const scale = toInt(pageSetup._$scale, settings.pageSetup.scale);
        settings.pageSetup.scale = clampScale(scale);
        settings.pageSetup.fitToWidth = toInt(pageSetup._$fitToWidth, settings.pageSetup.fitToWidth);
        settings.pageSetup.fitToHeight = toInt(pageSetup._$fitToHeight, settings.pageSetup.fitToHeight);

        const printOptions = xmlObj.printOptions || {};
        settings.printOptions.gridlines = toBool(printOptions._$gridLines ?? printOptions._$gridlines, settings.printOptions.gridlines);
        settings.printOptions.headings = toBool(printOptions._$headings, settings.printOptions.headings);

        const headerFooter = xmlObj.headerFooter || {};
        settings.headerFooter.header = headerFooter.oddHeader || '';
        settings.headerFooter.footer = headerFooter.oddFooter || '';

        const rowBreaks = normalizeArray(xmlObj.rowBreaks?.brk)
            .map(brk => toInt(brk?._$id, null))
            .filter(id => Number.isFinite(id))
            .map(id => id + 1);
        const colBreaks = normalizeArray(xmlObj.colBreaks?.brk)
            .map(brk => toInt(brk?._$id, null))
            .filter(id => Number.isFinite(id))
            .map(id => id + 1);
        settings.pageBreaks = { rowBreaks, colBreaks };

        if (sheet) {
            const printNames = this._readPrintDefinedNames(sheet);
            if (printNames.printArea) {
                settings.printArea = printNames.printArea;
            }
            if (printNames.printTitles.rows || printNames.printTitles.cols) {
                settings.printTitles = printNames.printTitles;
            }
        }

        return settings;
    }

    /**
     * Apply printing settings to worksheet XML
     * @param {Sheet} sheet - Sheet
     * @param {Object} ws - Worksheet XML Object
     * @returns {void}
     */
    applySettingsToWorksheet(sheet, ws) {
        const settings = this.ensureSettings(sheet);
        if (!settings || !ws) return;

        const margins = settings.pageMargins || DEFAULT_MARGINS;
        ws.pageMargins = {
            '_$left': String(margins.left ?? DEFAULT_MARGINS.left),
            '_$right': String(margins.right ?? DEFAULT_MARGINS.right),
            '_$top': String(margins.top ?? DEFAULT_MARGINS.top),
            '_$bottom': String(margins.bottom ?? DEFAULT_MARGINS.bottom),
            '_$header': String(margins.header ?? DEFAULT_MARGINS.header),
            '_$footer': String(margins.footer ?? DEFAULT_MARGINS.footer)
        };

        const setup = settings.pageSetup || {};
        ws.pageSetup = ws.pageSetup || {};
        ws.pageSetup._$paperSize = String(setup.paperSize || PAPER_SIZES.A4.paperSize);
        ws.pageSetup._$orientation = setup.orientation === 'landscape' ? 'landscape' : 'portrait';
        if (setup.scale) {
            ws.pageSetup._$scale = String(clampScale(setup.scale));
        } else if (ws.pageSetup._$scale) {
            delete ws.pageSetup._$scale;
        }
        if (setup.fitToWidth) ws.pageSetup._$fitToWidth = String(setup.fitToWidth);
        else if (ws.pageSetup._$fitToWidth) delete ws.pageSetup._$fitToWidth;
        if (setup.fitToHeight) ws.pageSetup._$fitToHeight = String(setup.fitToHeight);
        else if (ws.pageSetup._$fitToHeight) delete ws.pageSetup._$fitToHeight;

        const options = settings.printOptions || {};
        if (options.gridlines || options.headings) {
            ws.printOptions = ws.printOptions || {};
            if (options.gridlines) ws.printOptions._$gridLines = '1';
            else if (ws.printOptions._$gridLines) delete ws.printOptions._$gridLines;
            if (options.headings) ws.printOptions._$headings = '1';
            else if (ws.printOptions._$headings) delete ws.printOptions._$headings;
        } else if (ws.printOptions) {
            delete ws.printOptions;
        }

        const headerFooter = settings.headerFooter || {};
        if (headerFooter.header || headerFooter.footer) {
            ws.headerFooter = ws.headerFooter || {};
            if (headerFooter.header) ws.headerFooter.oddHeader = headerFooter.header;
            else if (ws.headerFooter.oddHeader) delete ws.headerFooter.oddHeader;
            if (headerFooter.footer) ws.headerFooter.oddFooter = headerFooter.footer;
            else if (ws.headerFooter.oddFooter) delete ws.headerFooter.oddFooter;
        } else if (ws.headerFooter) {
            delete ws.headerFooter;
        }

        const rowBreaks = settings.pageBreaks?.rowBreaks || [];
        const colBreaks = settings.pageBreaks?.colBreaks || [];
        if (rowBreaks.length > 0) {
            ws.rowBreaks = {
                '_$count': String(rowBreaks.length),
                '_$manualBreakCount': String(rowBreaks.length),
                brk: rowBreaks
                    .map(id => Math.max(0, toInt(id, 0) - 1))
                    .filter(id => Number.isFinite(id))
                    .map(id => ({ '_$id': String(id), '_$man': '1' }))
            };
        } else if (ws.rowBreaks) {
            delete ws.rowBreaks;
        }
        if (colBreaks.length > 0) {
            ws.colBreaks = {
                '_$count': String(colBreaks.length),
                '_$manualBreakCount': String(colBreaks.length),
                brk: colBreaks
                    .map(id => Math.max(0, toInt(id, 0) - 1))
                    .filter(id => Number.isFinite(id))
                    .map(id => ({ '_$id': String(id), '_$man': '1' }))
            };
        } else if (ws.colBreaks) {
            delete ws.colBreaks;
        }
    }

    /**
     * Get Print Settings (Deep Copy)
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    getSettings(sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        return settings ? clone(settings) : null;
    }

    /**
     * Merge Update Print Settings
     * @param {Object} partial - Partial Settings
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    setSettings(partial, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings || !partial) return settings;
        if (partial.pageSetup) Object.assign(settings.pageSetup, partial.pageSetup);
        if (partial.pageMargins) Object.assign(settings.pageMargins, partial.pageMargins);
        if (partial.printOptions) Object.assign(settings.printOptions, partial.printOptions);
        if (partial.printTitles) {
            settings.printTitles = this._normalizePrintTitles(partial.printTitles, sheet, settings.printTitles);
        }
        if (partial.headerFooter) Object.assign(settings.headerFooter, partial.headerFooter);
        if (partial.pageBreaks) Object.assign(settings.pageBreaks, partial.pageBreaks);
        if (partial.pageBackground !== undefined) settings.pageBackground = partial.pageBackground;
        if (partial.printArea !== undefined) settings.printArea = this.normalizeRange(partial.printArea, sheet);
        return settings;
    }

    /**
     * Set Page Margin
     * @param {Object} margins - Margin Configuration
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    setPageMargins(margins, sheet = this.SN.activeSheet) {
        return this.setSettings({ pageMargins: margins }, sheet);
    }

    /**
     * Set Page Configuration
     * @param {Object} setup - Page Settings
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    setPageSetup(setup, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return settings;
        if (setup.paperName) {
            const def = PAPER_SIZES[setup.paperName] || PAPER_SIZES.A4;
            settings.pageSetup.paperName = def.name;
            settings.pageSetup.paperSize = def.paperSize;
        }
        if (setup.paperSize && PAPER_SIZE_BY_ID[setup.paperSize]) {
            settings.pageSetup.paperSize = setup.paperSize;
            settings.pageSetup.paperName = PAPER_SIZE_BY_ID[setup.paperSize];
        }
        if (setup.orientation) settings.pageSetup.orientation = setup.orientation === 'landscape' ? 'landscape' : 'portrait';
        if (setup.scale !== undefined) settings.pageSetup.scale = clampScale(setup.scale);
        if (setup.fitToWidth !== undefined) settings.pageSetup.fitToWidth = toInt(setup.fitToWidth, 0);
        if (setup.fitToHeight !== undefined) settings.pageSetup.fitToHeight = toInt(setup.fitToHeight, 0);
        return settings;
    }

    /**
     * Set Print Options
     * @param {Object} options - Print Options
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    setPrintOptions(options, sheet = this.SN.activeSheet) {
        return this.setSettings({ printOptions: options }, sheet);
    }

    /**
     * Set Print Range
     * @param {string|Object} range - Regional
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    setPrintArea(range, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return null;
        settings.printArea = this.normalizeRange(range, sheet);
        return settings.printArea;
    }

    /**
     * Empty Print Area
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {void}
     */
    clearPrintArea(sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return;
        settings.printArea = null;
    }

    /**
     * Set the print title row/column
     * @param {rows? : stringing number[s]
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {void}
     */
    setPrintTitles(titles, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return;
        settings.printTitles = this._normalizePrintTitles(titles, sheet, settings.printTitles);
    }

    /**
     * Toggle Print Grid Line
     * @param {boolean} on - Enable
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {void}
     */
    toggleGridlines(on, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return;
        settings.printOptions.gridlines = !!on;
    }

    /**
     * Toggle Print Header
     * @param {boolean} on - Enable
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {void}
     */
    toggleHeadings(on, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return;
        settings.printOptions.headings = !!on;
    }

    /**
     * Show/Hide Page Break Guide
     * @param {boolean} on - Whether to show
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {void}
     */
    toggleShowPageBreaks(on, sheet = this.SN.activeSheet) {
        if (!sheet) return;
        sheet.showPageBreaks = !!on;
        if (sheet === this.SN.activeSheet) this.SN._r();
    }

    /**
     * Insert Manual Page Break in the current active cell
     * @param {r: number, c: number}
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {{row:boolean, col:boolean}|null}
     */
    insertPageBreak(cell = null, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings || !sheet) return null;
        const target = cell || sheet.activeCell;
        if (!target || !Number.isFinite(target.r) || !Number.isFinite(target.c)) return null;

        const rowBreaks = settings.pageBreaks?.rowBreaks || (settings.pageBreaks.rowBreaks = []);
        const colBreaks = settings.pageBreaks?.colBreaks || (settings.pageBreaks.colBreaks = []);
        let rowInserted = false;
        let colInserted = false;

        // 行分页符：在目标行上方分页，内部存储为“下一页起始行索引”
        if (target.r > 0 && !rowBreaks.includes(target.r)) {
            rowBreaks.push(target.r);
            rowBreaks.sort((a, b) => a - b);
            rowInserted = true;
        }
        // 列分页符：在目标列左侧分页，内部存储为“下一页起始列索引”
        if (target.c > 0 && !colBreaks.includes(target.c)) {
            colBreaks.push(target.c);
            colBreaks.sort((a, b) => a - b);
            colInserted = true;
        }

        return { row: rowInserted, col: colInserted };
    }

    /**
     * Get page break guide information (for canvas display)
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {{range:Object,rowBreaks:Array<{index:number,manual:boolean}>,colBreaks:Array<{index:number,manual:boolean}>}|null}
     */
    getPageBreakGuides(sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings || !sheet) return null;
        const layout = this.buildLayout(sheet, settings);
        const range = layout.range;
        if (!range) return null;

        const manualRowSet = new Set(
            (settings.pageBreaks?.rowBreaks || [])
                .map(val => toInt(val, null))
                .filter(val => Number.isFinite(val))
                .filter(val => val > range.s.r && val <= range.e.r)
        );
        const manualColSet = new Set(
            (settings.pageBreaks?.colBreaks || [])
                .map(val => toInt(val, null))
                .filter(val => Number.isFinite(val))
                .filter(val => val > range.s.c && val <= range.e.c)
        );

        const rowBreakSet = new Set(manualRowSet);
        const colBreakSet = new Set(manualColSet);
        layout.rowSegments.forEach((seg, idx) => {
            if (idx > 0 && seg && Number.isFinite(seg.s) && seg.s > range.s.r && seg.s <= range.e.r) {
                rowBreakSet.add(seg.s);
            }
        });
        layout.colSegments.forEach((seg, idx) => {
            if (idx > 0 && seg && Number.isFinite(seg.s) && seg.s > range.s.c && seg.s <= range.e.c) {
                colBreakSet.add(seg.s);
            }
        });

        return {
            range,
            rowBreaks: [...rowBreakSet].sort((a, b) => a - b).map(index => ({ index, manual: manualRowSet.has(index) })),
            colBreaks: [...colBreakSet].sort((a, b) => a - b).map(index => ({ index, manual: manualColSet.has(index) }))
        };
    }

    /**
     * Print Preview
     * @param {Object} [options={}] - Preview Options
     * @returns {Promise<void>}
     */
    async preview(options = {}) {
        const sheet = options.sheet || this.SN.activeSheet;
        if (!sheet) return;
        sheet._init();
        const settings = this.ensureSettings(sheet);
        const layout = this.buildLayout(sheet, settings, options);
        const pages = await this.renderPages(sheet, settings, layout);
        const html = this.buildPreviewHtml(pages, layout);

        await this.SN.Utils.modal({
            titleKey: 'core.print.print.modal.m001.title', title: 'Print Preview',
            content: html,
            maxWidth: '96vw',
            maxHeight: '90vh',
            confirmKey: 'core.print.print.modal.m001.confirm', confirmText: 'Print',
            cancelKey: 'core.print.print.modal.m001.cancel', cancelText: 'Close',
            bodyClass: 'sn-print-preview',
            onOpen: (bodyEl) => this.bindPreviewEvents(bodyEl, pages, layout),
            onConfirm: () => this.print({ sheet, settings, layout, pages })
        });
    }

    /**
     * Print
     * @param {Object} [options={}] - Print Options
     * @returns {Promise<void>}
     */
    async print(options = {}) {
        const sheet = options.sheet || this.SN.activeSheet;
        if (!sheet) return;
        sheet._init();
        const settings = options.settings || this.ensureSettings(sheet);
        const layout = options.layout || this.buildLayout(sheet, settings, options);
        const pages = options.pages || await this.renderPages(sheet, settings, layout);
        if (!pages.length) {
            this.SN.Utils.toast(this.SN.t('core.print.print.toast.nothingToPrint'));
            return;
        }

        const html = this.buildPrintHtml(pages, layout);
        const iframe = document.createElement('iframe');
        iframe.style.position = 'fixed';
        iframe.style.right = '0';
        iframe.style.bottom = '0';
        iframe.style.width = '0';
        iframe.style.height = '0';
        iframe.style.border = '0';

        iframe.onload = () => {
            const win = iframe.contentWindow;
            if (!win) return;
            win.focus();
            win.print();
            setTimeout(() => iframe.remove(), 1000);
        };

        iframe.srcdoc = html;
        this.SN.containerDom.appendChild(iframe);
    }

    /**
     * Build page break layout information
     * @param {Sheet} sheet - Sheet
     * @param {Object} settings - Print Settings
     * @returns {Object}
     */
    buildLayout(sheet, settings, options = {}) {
        const paper = this.getPaperDefinition(settings);
        const margins = this.getMarginPixels(settings.pageMargins);
        const printable = {
            width: Math.max(0, paper.widthPx - margins.left - margins.right),
            height: Math.max(0, paper.heightPx - margins.top - margins.bottom)
        };
        const scale = this._resolvePrintScale(sheet, settings, printable);
        const contentLimit = {
            width: printable.width / scale,
            height: printable.height / scale
        };
        const range = this.getPrintRange(sheet, settings);
        const repeatTitles = this._resolveRepeatTitles(sheet, settings, range);
        const bodyRange = this._buildBodyRange(range, repeatTitles);
        const bodyLimit = {
            width: Math.max(1, contentLimit.width - repeatTitles.colsSize),
            height: Math.max(1, contentLimit.height - repeatTitles.rowsSize)
        };

        let rowSegments = [];
        let colSegments = [];
        if (bodyRange) {
            rowSegments = this.buildSegments(
                bodyRange.s.r,
                bodyRange.e.r,
                (r) => this.getRowSize(sheet, r),
                bodyLimit.height,
                settings.pageBreaks?.rowBreaks || []
            );
            colSegments = this.buildSegments(
                bodyRange.s.c,
                bodyRange.e.c,
                (c) => this.getColSize(sheet, c),
                bodyLimit.width,
                settings.pageBreaks?.colBreaks || []
            );
        } else {
            rowSegments = [null];
            colSegments = [null];
        }

        if (!rowSegments.length) rowSegments = [null];
        if (!colSegments.length) colSegments = [null];

        return {
            paper,
            margins,
            printable,
            scale,
            range,
            bodyRange,
            rowSegments,
            colSegments,
            repeatTitles,
            pageBreakPreview: !!options.pageBreakPreview,
            settings
        };
    }

    /**
     * Access to paper definition
     * @param {Object} settings - Print Settings
     * @returns {Object}
     */
    getPaperDefinition(settings) {
        const setup = settings?.pageSetup || {};
        const paperName = setup.paperName || PAPER_SIZE_BY_ID[setup.paperSize] || 'A4';
        const base = PAPER_SIZES[paperName] || PAPER_SIZES.A4;
        let widthIn = base.widthIn;
        let heightIn = base.heightIn;
        if (setup.orientation === 'landscape') {
            [widthIn, heightIn] = [heightIn, widthIn];
        }
        return {
            name: base.name,
            paperSize: base.paperSize,
            widthIn,
            heightIn,
            widthPx: Math.round(widthIn * PRINT_DPI),
            heightPx: Math.round(heightIn * PRINT_DPI)
        };
    }

    /**
     * Calculate margin pixels
     * @param {Object} [margins=DEFAULT_MARGINS] - Margin Configuration
     * @returns {{left: number, right: number, top: number, bottom: number}}
     */
    getMarginPixels(margins = DEFAULT_MARGINS) {
        return {
            left: Math.round((margins.left ?? DEFAULT_MARGINS.left) * PRINT_DPI),
            right: Math.round((margins.right ?? DEFAULT_MARGINS.right) * PRINT_DPI),
            top: Math.round((margins.top ?? DEFAULT_MARGINS.top) * PRINT_DPI),
            bottom: Math.round((margins.bottom ?? DEFAULT_MARGINS.bottom) * PRINT_DPI)
        };
    }

    /**
     * Get Actual Print Range
     * @param {Sheet} sheet - Sheet
     * @param {Object} settings - Print Settings
     * @returns {{s: {r: number, c: number}, e: {r: number, c: number}}}
     */
    getPrintRange(sheet, settings) {
        const area = settings?.printArea;
        const range = this.normalizeRange(area, sheet);
        if (range) return range;
        return this.getUsedRange(sheet);
    }

    /**
     * Normative Scope Object
     * @param {string|Object|null} range - Scope
     * @param {Sheet} sheet - Sheet
     * @returns {{s: {r: number, c: number}, e: {r: number, c: number}}|null}
     */
    normalizeRange(range, sheet) {
        if (!range || !sheet) return null;
        if (typeof range === 'string') return sheet.rangeStrToNum(range);
        if (range.s && range.e) return range;
        return null;
    }

    /**
     * Get used area
     * @param {Sheet} sheet - Sheet
     * @returns {{s: {r: number, c: number}, e: {r: number, c: number}}}
     */
    getUsedRange(sheet) {
        let minR = Infinity;
        let minC = Infinity;
        let maxR = -1;
        let maxC = -1;

        sheet.rows.forEach((row, rIndex) => {
            if (!row) return;
            row.cells.forEach((cell, cIndex) => {
                if (!cell || !cell._needBuild) return;
                minR = Math.min(minR, rIndex);
                minC = Math.min(minC, cIndex);
                maxR = Math.max(maxR, rIndex);
                maxC = Math.max(maxC, cIndex);
            });
        });

        if (maxR < 0 || maxC < 0) {
            return { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
        }
        return { s: { r: minR, c: minC }, e: { r: maxR, c: maxC } };
    }

    /**
     * Build page segment
     * @param {number} start - Initial Index
     * @param {number} end - End Index
     * @param {Function} sizeGetter - Fetch Size Method
     * @param {number} maxSize - Maximum size
     * @returns {Array<{s: number, e: number}>}
     */
    buildSegments(start, end, sizeGetter, maxSize, forcedStarts = []) {
        if (!Number.isFinite(start) || !Number.isFinite(end) || start > end) return [];
        const segments = [];
        const maxSegmentSize = Math.max(1, toNumber(maxSize, 1));
        const forcedSet = new Set(
            (forcedStarts || [])
                .map(val => toInt(val, null))
                .filter(val => Number.isFinite(val) && val > start && val <= end)
        );
        let segStart = start;
        let size = 0;
        for (let i = start; i <= end; i++) {
            if (forcedSet.has(i) && i > segStart) {
                segments.push({ s: segStart, e: i - 1 });
                segStart = i;
                size = 0;
            }
            const nextSize = Math.max(0, sizeGetter(i));
            if (size + nextSize > maxSegmentSize && i > segStart) {
                segments.push({ s: segStart, e: i - 1 });
                segStart = i;
                size = 0;
            }
            size += nextSize;
        }
        if (segStart <= end) {
            segments.push({ s: segStart, e: end });
        }
        return segments;
    }

    /**
     * Get line height (pixels)
     * @param {Sheet} sheet - Sheet
     * @param {number} r - Row index
     * @returns {number}
     */
    getRowSize(sheet, r) {
        const row = sheet.rows[r];
        if (row && row.hidden) return 0;
        return row ? row.height : sheet.defaultRowHeight;
    }

    /**
     * Get column width (pixels)
     * @param {Sheet} sheet - Sheet
     * @param {number} c - Column Index
     * @returns {number}
     */
    getColSize(sheet, c) {
        const col = sheet.cols[c];
        if (col && col.hidden) return 0;
        return col ? col.width : sheet.defaultColWidth;
    }

    /**
     * Get range dimensions
     * @param {Sheet} sheet - Sheet
     * @param {s: {r: number, c: number}, e: {r: number, c: number}}range - range
     * @param {boolean} includeHeadings - Whether to include a column header
     * @returns {{width: number, height: number}}
     */
    getRangeSize(sheet, range, includeHeadings) {
        let width = 0;
        let height = 0;
        for (let c = range.s.c; c <= range.e.c; c++) {
            width += this.getColSize(sheet, c);
        }
        for (let r = range.s.r; r <= range.e.r; r++) {
            height += this.getRowSize(sheet, r);
        }
        if (includeHeadings) {
            width += sheet._indexWidth ?? sheet.indexWidth;
            height += sheet._headHeight ?? sheet.headHeight;
        }
        return { width, height };
    }

    /**
     * Build cell address with sheet name
     * @param {Sheet} sheet - Sheet
     * @param {r: number, c: number}cell - cell coordinates
     * @returns {string}
     */
    buildCellAddress(sheet, cell) {
        const ref = this.SN.Utils.cellNumToStr(cell);
        const name = sheet.name || '';
        if (!name) return ref;
        if (/[\\s!]/.test(name)) return `'${name}'!${ref}`;
        return `${name}!${ref}`;
    }

    /**
     * Render all pages
     * @param {Sheet} sheet - Sheet
     * @param {Object} settings - Print Settings
     * @param {Object} layout - Layout Information
     * @returns {Promise<Array<Object>>}
     */
    async renderPages(sheet, settings, layout) {
        const pages = [];
        const includeHeadings = settings.printOptions?.headings;
        const rowSegments = layout.rowSegments?.length ? layout.rowSegments : [null];
        const colSegments = layout.colSegments?.length ? layout.colSegments : [null];
        for (const rowSeg of rowSegments) {
            for (const colSeg of colSegments) {
                const range = this._buildPageRange(layout, rowSeg, colSeg);
                const size = this.getRangeSize(sheet, range, includeHeadings);
                const dataUrl = await this.captureRange(sheet, range, size.width, size.height, settings);
                pages.push({
                    range,
                    rowSeg,
                    colSeg,
                    width: size.width,
                    height: size.height,
                    dataUrl
                });
            }
        }
        return pages;
    }

    /**
     * Caption range specified
     * @param {Sheet} sheet - Sheet
     * @param {s: {r: number, c: number}, e: {r: number, c: number}}range - range
     * @param {number} width - Width
     * @param {number} height - Height
     * @param {Object} settings - Print Settings
     * @returns {Promise<string>}
     */
    async captureRange(sheet, range, width, height, settings) {
        const prevGrid = sheet.showGridLines;
        const prevHeadings = sheet.showRowColHeaders;
        sheet.showGridLines = !!settings.printOptions?.gridlines;
        sheet.showRowColHeaders = !!settings.printOptions?.headings;

        try {
            const safeWidth = Math.max(1, Math.round(width || 1));
            const safeHeight = Math.max(1, Math.round(height || 1));
            const cellAddress = this.buildCellAddress(sheet, range.s);
            return await this.SN.Canvas.captureScreenshot(cellAddress, 'topleft', {
                width: safeWidth,
                height: safeHeight,
                compress: false,
                hideSelectionFrame: true
            });
        } finally {
            sheet.showGridLines = prevGrid;
            sheet.showRowColHeaders = prevHeadings;
        }
    }

    _resolvePrintScale(sheet, settings, printable) {
        const setup = settings.pageSetup || {};
        let scale = (setup.scale ?? 100) / 100;
        const fitToWidth = Math.max(0, toInt(setup.fitToWidth, 0));
        const fitToHeight = Math.max(0, toInt(setup.fitToHeight, 0));
        if (fitToWidth <= 0 && fitToHeight <= 0) {
            return Math.max(0.1, Math.min(4, scale || 1));
        }
        const range = this.getPrintRange(sheet, settings);
        const size = this.getRangeSize(sheet, range, !!settings.printOptions?.headings);
        const byWidth = fitToWidth > 0 ? (printable.width * fitToWidth) / Math.max(1, size.width) : Infinity;
        const byHeight = fitToHeight > 0 ? (printable.height * fitToHeight) / Math.max(1, size.height) : Infinity;
        const fitScale = Math.min(byWidth, byHeight);
        if (Number.isFinite(fitScale) && fitScale > 0) {
            scale = fitScale;
        }
        return Math.max(0.1, Math.min(4, scale || 1));
    }

    _resolveRepeatTitles(sheet, settings, range) {
        const rows = this._parseTitleRows(settings.printTitles?.rows);
        const cols = this._parseTitleCols(settings.printTitles?.cols);
        let repeatRows = null;
        let repeatCols = null;

        if (rows) {
            const next = {
                s: Math.max(range.s.r, rows.s),
                e: Math.min(range.e.r, rows.e)
            };
            if (next.s <= next.e && next.s === range.s.r) repeatRows = next;
        }
        if (cols) {
            const next = {
                s: Math.max(range.s.c, cols.s),
                e: Math.min(range.e.c, cols.e)
            };
            if (next.s <= next.e && next.s === range.s.c) repeatCols = next;
        }

        let rowsSize = 0;
        let colsSize = 0;
        if (repeatRows) {
            for (let r = repeatRows.s; r <= repeatRows.e; r++) rowsSize += this.getRowSize(sheet, r);
        }
        if (repeatCols) {
            for (let c = repeatCols.s; c <= repeatCols.e; c++) colsSize += this.getColSize(sheet, c);
        }

        return {
            rows: repeatRows,
            cols: repeatCols,
            rowsSize,
            colsSize
        };
    }

    _buildBodyRange(range, repeatTitles) {
        const startRow = repeatTitles.rows ? repeatTitles.rows.e + 1 : range.s.r;
        const startCol = repeatTitles.cols ? repeatTitles.cols.e + 1 : range.s.c;
        if (startRow > range.e.r || startCol > range.e.c) return null;
        return {
            s: { r: startRow, c: startCol },
            e: { r: range.e.r, c: range.e.c }
        };
    }

    _buildPageRange(layout, rowSeg, colSeg) {
        const repeatRows = layout.repeatTitles?.rows || null;
        const repeatCols = layout.repeatTitles?.cols || null;
        const bodyRows = rowSeg || repeatRows || { s: layout.range.s.r, e: layout.range.e.r };
        const bodyCols = colSeg || repeatCols || { s: layout.range.s.c, e: layout.range.e.c };
        return {
            s: {
                r: repeatRows ? repeatRows.s : bodyRows.s,
                c: repeatCols ? repeatCols.s : bodyCols.s
            },
            e: {
                r: bodyRows.e,
                c: bodyCols.e
            }
        };
    }

    _normalizePrintTitles(titles, sheet, prev = { rows: null, cols: null }) {
        const safeTitles = titles || {};
        const hasRows = Object.prototype.hasOwnProperty.call(safeTitles, 'rows');
        const hasCols = Object.prototype.hasOwnProperty.call(safeTitles, 'cols');
        const normalizedRows = this._normalizeTitleRowsValue(safeTitles.rows, sheet);
        const normalizedCols = this._normalizeTitleColsValue(safeTitles.cols, sheet);
        return {
            rows: hasRows
                ? (safeTitles.rows == null || safeTitles.rows === '' ? null : (normalizedRows ?? (prev?.rows ?? null)))
                : (prev?.rows ?? null),
            cols: hasCols
                ? (safeTitles.cols == null || safeTitles.cols === '' ? null : (normalizedCols ?? (prev?.cols ?? null)))
                : (prev?.cols ?? null)
        };
    }

    _normalizeTitleRowsValue(value, sheet) {
        if (value == null) return null;
        if (Array.isArray(value)) {
            const nums = value.map(v => toInt(v, null)).filter(v => Number.isFinite(v));
            if (!nums.length) return null;
            const min = Math.min(...nums);
            const max = Math.max(...nums);
            return `${min + 1}:${max + 1}`;
        }
        if (typeof value === 'object' && Number.isFinite(value.s) && Number.isFinite(value.e)) {
            const min = Math.min(value.s, value.e);
            const max = Math.max(value.s, value.e);
            return `${min + 1}:${max + 1}`;
        }
        if (typeof value !== 'string') return null;
        const ref = value.trim();
        if (!ref) return null;
        const part = this._stripSheetPrefix(ref, sheet?.name);
        const clean = part.replace(/\$/g, '').replace(/\s+/g, '');
        if (!/^\d+:\d+$/.test(clean)) return null;
        const [sText, eText] = clean.split(':');
        const s = Math.max(1, toInt(sText, 1));
        const e = Math.max(1, toInt(eText, s));
        return `${Math.min(s, e)}:${Math.max(s, e)}`;
    }

    _normalizeTitleColsValue(value, sheet) {
        if (value == null) return null;
        if (Array.isArray(value)) {
            const nums = value.map(v => toInt(v, null)).filter(v => Number.isFinite(v));
            if (!nums.length) return null;
            const min = Math.min(...nums);
            const max = Math.max(...nums);
            return `${this.SN.Utils.numToChar(min)}:${this.SN.Utils.numToChar(max)}`;
        }
        if (typeof value === 'object' && Number.isFinite(value.s) && Number.isFinite(value.e)) {
            const min = Math.min(value.s, value.e);
            const max = Math.max(value.s, value.e);
            return `${this.SN.Utils.numToChar(min)}:${this.SN.Utils.numToChar(max)}`;
        }
        if (typeof value !== 'string') return null;
        const ref = value.trim();
        if (!ref) return null;
        const part = this._stripSheetPrefix(ref, sheet?.name);
        const clean = part.replace(/\$/g, '').replace(/\s+/g, '').toUpperCase();
        if (!/^[A-Z]+:[A-Z]+$/.test(clean)) return null;
        const [sText, eText] = clean.split(':');
        const s = this.SN.Utils.charToNum(sText);
        const e = this.SN.Utils.charToNum(eText);
        return `${this.SN.Utils.numToChar(Math.min(s, e))}:${this.SN.Utils.numToChar(Math.max(s, e))}`;
    }

    _parseTitleRows(rowsRef) {
        if (!rowsRef || typeof rowsRef !== 'string') return null;
        const match = rowsRef.trim().match(/^(\d+):(\d+)$/);
        if (!match) return null;
        const s = Math.max(0, toInt(match[1], 1) - 1);
        const e = Math.max(0, toInt(match[2], 1) - 1);
        return { s: Math.min(s, e), e: Math.max(s, e) };
    }

    _parseTitleCols(colsRef) {
        if (!colsRef || typeof colsRef !== 'string') return null;
        const match = colsRef.trim().toUpperCase().match(/^([A-Z]+):([A-Z]+)$/);
        if (!match) return null;
        const s = this.SN.Utils.charToNum(match[1]);
        const e = this.SN.Utils.charToNum(match[2]);
        return { s: Math.min(s, e), e: Math.max(s, e) };
    }

    _readPrintDefinedNames(sheet) {
        const out = {
            printArea: null,
            printTitles: { rows: null, cols: null }
        };
        const list = this._getWorkbookDefinedNameList();
        if (!list.length || !sheet) return out;

        const sheetIndex = this.SN.sheets.indexOf(sheet);
        const printAreaItem = list.find(item => item?._$name === '_xlnm.Print_Area' && this._isDefinedNameForSheet(item, sheet, sheetIndex));
        if (printAreaItem?.['#text']) {
            const areaPart = this._splitDefinedRefList(printAreaItem['#text'])
                .map(part => this._stripSheetPrefix(part, sheet.name))
                .find(part => !!part);
            if (areaPart) {
                const normalizedArea = this.normalizeRange(areaPart, sheet);
                out.printArea = normalizedArea || areaPart;
            }
        }

        const printTitlesItem = list.find(item => item?._$name === '_xlnm.Print_Titles' && this._isDefinedNameForSheet(item, sheet, sheetIndex));
        if (printTitlesItem?.['#text']) {
            const refs = this._splitDefinedRefList(printTitlesItem['#text'])
                .map(part => this._stripSheetPrefix(part, sheet.name))
                .filter(Boolean);
            refs.forEach(ref => {
                const rows = this._normalizeTitleRowsValue(ref, sheet);
                const cols = this._normalizeTitleColsValue(ref, sheet);
                if (rows) out.printTitles.rows = rows;
                if (cols) out.printTitles.cols = cols;
            });
        }

        return out;
    }

    _getWorkbookDefinedNameList() {
        const defined = this.SN.Xml?.obj?.['xl/workbook.xml']?.workbook?.definedNames?.definedName;
        if (!defined) return [];
        return Array.isArray(defined) ? defined : [defined];
    }

    _isDefinedNameForSheet(item, sheet, sheetIndex) {
        if (!item || !sheet) return false;
        if (item._$localSheetId !== undefined) {
            return toInt(item._$localSheetId, -1) === sheetIndex;
        }
        const text = item['#text'];
        if (typeof text !== 'string' || !text.includes('!')) return false;
        return this._splitDefinedRefList(text).some(part => {
            const parsed = this._parseQualifiedRef(part);
            return parsed.sheetName === sheet.name;
        });
    }

    _splitDefinedRefList(text) {
        if (!text || typeof text !== 'string') return [];
        const parts = [];
        let current = '';
        let inQuotes = false;
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (ch === "'") {
                if (inQuotes && text[i + 1] === "'") {
                    current += "''";
                    i += 1;
                    continue;
                }
                inQuotes = !inQuotes;
                current += ch;
                continue;
            }
            if (ch === ',' && !inQuotes) {
                if (current.trim()) parts.push(current.trim());
                current = '';
                continue;
            }
            current += ch;
        }
        if (current.trim()) parts.push(current.trim());
        return parts;
    }

    _parseQualifiedRef(refText) {
        const text = String(refText || '').trim();
        const match = text.match(/^('(?:[^']|'')+'|[^'!]+)!(.+)$/);
        if (!match) return { sheetName: null, ref: text };
        const rawSheet = match[1];
        const sheetName = rawSheet.startsWith("'")
            ? rawSheet.slice(1, -1).replace(/''/g, "'")
            : rawSheet;
        return { sheetName, ref: match[2] };
    }

    _stripSheetPrefix(refText, expectedSheetName = null) {
        const parsed = this._parseQualifiedRef(refText);
        if (!parsed.sheetName) return parsed.ref;
        if (expectedSheetName && parsed.sheetName !== expectedSheetName) return '';
        return parsed.ref;
    }

    _resolveHeaderFooterText(template, context = {}) {
        const raw = String(template ?? '');
        if (!raw) return '';
        const now = new Date();
        const dateText = context.dateText ?? now.toLocaleDateString();
        const timeText = context.timeText ?? now.toLocaleTimeString();
        const pageNumber = context.pageNumber ?? '';
        const totalPages = context.totalPages ?? '';
        const sheetName = context.sheetName ?? '';

        return raw
            .replace(/&&/g, '\u0000')
            .replace(/&[LCR]/gi, '')
            .replace(/&P/gi, String(pageNumber))
            .replace(/&N/gi, String(totalPages))
            .replace(/&A/gi, String(sheetName))
            .replace(/&D/gi, String(dateText))
            .replace(/&T/gi, String(timeText))
            .replace(/&[FZG]/gi, '')
            .replace(/&\"[^\"]*\"/g, '')
            .replace(/&K[0-9A-Fa-f]{1,8}/g, '')
            .replace(/&\d+/g, '')
            .replace(/&[BIESUXYOH]/gi, '')
            .replace(/&[+-]/g, '')
            .replace(/\u0000/g, '&')
            .trim();
    }

    _buildPageStyle(settings, forAttr = true) {
        const bgUrl = settings.pageBackground?.url;
        if (!bgUrl || typeof bgUrl !== 'string') return '';
        const trimmed = bgUrl.trim();
        if (!trimmed) return '';
        if (/^\s*javascript:/i.test(trimmed)) return '';
        const safeUrl = forAttr
            ? escapeAttr(trimmed)
            : trimmed.replace(/["\\]/g, '\\$&');
        return `background-image:url("${safeUrl}");background-repeat:no-repeat;background-position:center center;background-size:cover;`;
    }

    _buildPageHeaderFooterHtml(settings, margins, paper, context = {}) {
        const header = this._resolveHeaderFooterText(settings.headerFooter?.header ?? '', context);
        const footer = this._resolveHeaderFooterText(settings.headerFooter?.footer ?? '', context);
        const headerText = escapeHtml(header).replace(/\n/g, '<br>');
        const footerText = escapeHtml(footer).replace(/\n/g, '<br>');
        const headerPx = Math.round(((settings.pageMargins?.header ?? 0.3) * PRINT_DPI));
        const footerPx = Math.round(((settings.pageMargins?.footer ?? 0.3) * PRINT_DPI));
        const baseStyle = `left:${margins.left}px;right:${margins.right}px;max-width:${paper.widthPx - margins.left - margins.right}px;`;
        return {
            header: headerText
                ? `<div class="sn-print-page-header" style="${baseStyle}top:${Math.max(4, headerPx - 12)}px;">${headerText}</div>`
                : '',
            footer: footerText
                ? `<div class="sn-print-page-footer" style="${baseStyle}bottom:${Math.max(4, footerPx - 12)}px;">${footerText}</div>`
                : ''
        };
    }

    /**
     * Build Preview HTML
     * @param {Array<Object>} pages - Page Data
     * @param {Object} layout - Layout Information
     * @returns {string}
     */
    buildPreviewHtml(pages, layout) {
        if (!pages.length) {
            return `<div class="sn-print-preview-empty">Nothing to print</div>`;
        }
        const thumbs = pages.map((page, idx) => (
            `<div class="sn-print-thumb" data-index="${idx}">
                <div class="sn-print-thumb-label">Page ${idx + 1}</div>
                <img src="${page.dataUrl}" alt="page-${idx + 1}">
            </div>`
        )).join('');

        const main = this.buildPreviewPageHtml(pages[0], layout, 0, pages.length);

        return `
            <div class="sn-print-preview-sidebar">${thumbs}</div>
            <div class="sn-print-preview-main">${main}</div>
        `;
    }

    /**
     * Build Single Page Preview HTML
     * @param {Object} page - Page Data
     * @param {Object} layout - Layout Information
     * @returns {string}
     */
    buildPreviewPageHtml(page, layout, pageIndex = 0, pageCount = 1) {
        const { paper, margins, scale, settings } = layout;
        const imgWidth = Math.round(page.width * scale);
        const imgHeight = Math.round(page.height * scale);
        const pageStyle = this._buildPageStyle(settings, true);
        const headerFooterHtml = this._buildPageHeaderFooterHtml(settings, margins, paper, {
            pageNumber: pageIndex + 1,
            totalPages: pageCount,
            sheetName: this.SN.activeSheet?.name || ''
        });
        const rangeBadge = layout.pageBreakPreview
            ? `<div class="sn-print-page-range">Page ${pageIndex + 1} · ${this.SN.Utils.rangeNumToStr(page.range)}</div>`
            : '';
        return `
            <div class="sn-print-page" style="width:${paper.widthPx}px;height:${paper.heightPx}px;padding:${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px;${pageStyle}">
                ${headerFooterHtml.header}
                ${rangeBadge}
                <img class="sn-print-page-image" src="${page.dataUrl}" style="width:${imgWidth}px;height:${imgHeight}px;" alt="print-page">
                ${headerFooterHtml.footer}
            </div>
        `;
    }

    /**
     * Tie preview interactive events
     * @param {HTMLElement} bodyEl - Pop the window, buddy.
     * @param {Array<Object>} pages - Page Data
     * @param {Object} layout - Layout Information
     * @returns {void}
     */
    bindPreviewEvents(bodyEl, pages, layout) {
        if (!pages.length) return;
        const main = bodyEl.querySelector('.sn-print-preview-main');
        const thumbs = Array.from(bodyEl.querySelectorAll('.sn-print-thumb'));
        const renderMain = (index) => {
            const page = pages[index];
            main.innerHTML = this.buildPreviewPageHtml(page, layout, index, pages.length);
            thumbs.forEach((item, i) => item.classList.toggle('active', i === index));
        };
        thumbs.forEach((item, index) => {
            item.addEventListener('click', () => renderMain(index));
        });
        renderMain(0);
    }

    /**
     * Build HTML
     * @param {Array<Object>} pages - Page Data
     * @param {Object} layout - Layout Information
     * @returns {string}
     */
    buildPrintHtml(pages, layout) {
        const { paper, margins, scale, settings } = layout;
        const pageCssSize = `${paper.name} ${paper.widthIn > paper.heightIn ? 'landscape' : 'portrait'}`;
        const headerPx = Math.round(((settings.pageMargins?.header ?? 0.3) * PRINT_DPI));
        const footerPx = Math.round(((settings.pageMargins?.footer ?? 0.3) * PRINT_DPI));
        const pageStyles = `
            @page { size: ${pageCssSize}; margin: 0; }
            body { margin: 0; padding: 0; }
            .page {
                position: relative;
                width: ${paper.widthPx}px;
                height: ${paper.heightPx}px;
                padding: ${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px;
                box-sizing: border-box;
                page-break-after: always;
                ${this._buildPageStyle(settings, false)}
            }
            .page img { display: block; }
            .page .sn-print-page-header,
            .page .sn-print-page-footer {
                position: absolute;
                left: ${margins.left}px;
                right: ${margins.right}px;
                color: #6b7280;
                font-size: 12px;
                white-space: pre-wrap;
                pointer-events: none;
            }
            .page .sn-print-page-header { top: ${Math.max(4, headerPx - 12)}px; }
            .page .sn-print-page-footer { bottom: ${Math.max(4, footerPx - 12)}px; }
        `;
        const pagesHtml = pages.map((page, pageIndex) => {
            const imgWidth = Math.round(page.width * scale);
            const imgHeight = Math.round(page.height * scale);
            const headerFooterHtml = this._buildPageHeaderFooterHtml(settings, margins, paper, {
                pageNumber: pageIndex + 1,
                totalPages: pages.length,
                sheetName: this.SN.activeSheet?.name || ''
            });
            return `
                <div class="page">
                    ${headerFooterHtml.header}
                    <img src="${page.dataUrl}" style="width:${imgWidth}px;height:${imgHeight}px;">
                    ${headerFooterHtml.footer}
                </div>
            `;
        }).join('');

        return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>${pageStyles}</style>
</head>
<body>${pagesHtml}</body>
</html>`;
    }
}
