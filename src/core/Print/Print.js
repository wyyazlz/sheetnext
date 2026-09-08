import PrintDialog from '../../components/PrintDialog.js';

const PRINT_DPI = 96;

const PAPER_SIZES = {
    A4: { name: 'A4', paperSize: 9, widthIn: 210 / 25.4, heightIn: 297 / 25.4 },
    A3: { name: 'A3', paperSize: 8, widthIn: 297 / 25.4, heightIn: 420 / 25.4 },
    A5: { name: 'A5', paperSize: 11, widthIn: 148 / 25.4, heightIn: 210 / 25.4 },
    Legal: { name: 'Legal', paperSize: 5, widthIn: 8.5, heightIn: 14 },
    Letter: { name: 'Letter', paperSize: 1, widthIn: 8.5, heightIn: 11 }
};

const PAPER_SIZE_BY_ID = {
    9: 'A4',
    8: 'A3',
    11: 'A5',
    5: 'Legal',
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

const MARGIN_PRESETS = {
    normal: DEFAULT_MARGINS,
    narrow: { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25, header: 0.2, footer: 0.2 },
    wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.3, footer: 0.3 }
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
        this._dialog = null;
        this._settingsRefreshPending = false;
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
        if (!sheet.initialized) sheet._init();
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
        const fitToPage = toBool(xmlObj.sheetPr?.pageSetUpPr?._$fitToPage);
        settings.pageSetup.fitToWidth = fitToPage ? toInt(pageSetup._$fitToWidth, 1) : 0;
        settings.pageSetup.fitToHeight = fitToPage ? toInt(pageSetup._$fitToHeight, 1) : 0;

        const printOptions = xmlObj.printOptions || {};
        settings.printOptions.gridlines = toBool(printOptions._$gridLines ?? printOptions._$gridlines, settings.printOptions.gridlines);
        settings.printOptions.headings = toBool(printOptions._$headings, settings.printOptions.headings);

        const headerFooter = xmlObj.headerFooter || {};
        settings.headerFooter.header = headerFooter.oddHeader || '';
        settings.headerFooter.footer = headerFooter.oddFooter || '';

        // OOXML uses the one-based row/column BEFORE the break, which equals
        // our zero-based start index of the next page. Automatic breaks are recalculated.
        const rowBreaks = this._normalizePageBreaks(normalizeArray(xmlObj.rowBreaks?.brk)
            .filter(brk => toBool(brk?._$man))
            .map(brk => brk?._$id), 1048576);
        const colBreaks = this._normalizePageBreaks(normalizeArray(xmlObj.colBreaks?.brk)
            .filter(brk => toBool(brk?._$man))
            .map(brk => brk?._$id), 16384);
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
        const fitToPage = setup.fitToWidth > 0 || setup.fitToHeight > 0;
        if (fitToPage) {
            ws.pageSetup._$fitToWidth = String(setup.fitToWidth || 0);
            ws.pageSetup._$fitToHeight = String(setup.fitToHeight || 0);
            ws.sheetPr = ws.sheetPr || {};
            ws.sheetPr.pageSetUpPr = { ...ws.sheetPr.pageSetUpPr, '_$fitToPage': '1' };
        } else {
            delete ws.pageSetup._$fitToWidth;
            delete ws.pageSetup._$fitToHeight;
            if (ws.sheetPr?.pageSetUpPr) delete ws.sheetPr.pageSetUpPr._$fitToPage;
        }

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

        const rowBreaks = this._normalizePageBreaks(settings.pageBreaks?.rowBreaks, 1048576);
        const colBreaks = this._normalizePageBreaks(settings.pageBreaks?.colBreaks, 16384);
        if (rowBreaks.length > 0) {
            ws.rowBreaks = {
                '_$count': String(rowBreaks.length),
                '_$manualBreakCount': String(rowBreaks.length),
                brk: rowBreaks.map(id => ({ '_$id': String(id), '_$min': '0', '_$max': '16383', '_$man': '1' }))
            };
        } else if (ws.rowBreaks) {
            delete ws.rowBreaks;
        }
        if (colBreaks.length > 0) {
            ws.colBreaks = {
                '_$count': String(colBreaks.length),
                '_$manualBreakCount': String(colBreaks.length),
                brk: colBreaks.map(id => ({ '_$id': String(id), '_$min': '0', '_$max': '1048575', '_$man': '1' }))
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
        if (partial.pageBreaks) {
            if (partial.pageBreaks.rowBreaks !== undefined) settings.pageBreaks.rowBreaks = this._normalizePageBreaks(partial.pageBreaks.rowBreaks, 1048576);
            if (partial.pageBreaks.colBreaks !== undefined) settings.pageBreaks.colBreaks = this._normalizePageBreaks(partial.pageBreaks.colBreaks, 16384);
        }
        if (partial.pageBackground !== undefined) settings.pageBackground = partial.pageBackground;
        if (partial.printArea !== undefined) settings.printArea = this.normalizeRange(partial.printArea, sheet);
        this._notifySettingsChanged(sheet);
        return settings;
    }

    /**
     * Set Page Margin
     * @param {Object} margins - Margin Configuration
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {Object|null}
     */
    setPageMargins(margins, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return null;
        const next = { ...settings.pageMargins, ...margins };
        const paper = this.getPaperDefinition(settings);
        if (Object.values(next).some(value => !Number.isFinite(Number(value)) || Number(value) < 0)
            || Number(next.left) + Number(next.right) >= paper.widthIn
            || Number(next.top) + Number(next.bottom) >= paper.heightIn) {
            throw new Error(this.SN.t('core.print.print.invalidMargins'));
        }
        return this.setSettings({ pageMargins: margins }, sheet);
    }

    /** @returns {Array<Object>} Supported paper sizes in inches. */
    getPaperSizes() {
        return Object.values(PAPER_SIZES).map(paper => ({ ...paper }));
    }

    /** @returns {Object} Named margin presets in inches. */
    getMarginPresets() {
        return clone(MARGIN_PRESETS);
    }

    /** @param {Sheet} [sheet=this.SN.activeSheet] @returns {'normal'|'narrow'|'wide'|'custom'} Matching margin preset. */
    getPageMarginsPreset(sheet = this.SN.activeSheet) {
        const margins = this.ensureSettings(sheet)?.pageMargins;
        return Object.keys(MARGIN_PRESETS).find(name => Object.entries(MARGIN_PRESETS[name])
            .every(([field, value]) => Math.abs(margins?.[field] - value) < 0.001)) || 'custom';
    }

    /** @param {string} preset @param {Sheet} [sheet=this.SN.activeSheet] */
    setPageMarginsPreset(preset, sheet = this.SN.activeSheet) {
        return this.setPageMargins(MARGIN_PRESETS[preset] || DEFAULT_MARGINS, sheet);
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
        if (setup.fitToWidth !== undefined) settings.pageSetup.fitToWidth = Math.max(0, toInt(setup.fitToWidth, 0));
        if (setup.fitToHeight !== undefined) settings.pageSetup.fitToHeight = Math.max(0, toInt(setup.fitToHeight, 0));
        this._notifySettingsChanged(sheet);
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
        this._notifySettingsChanged(sheet);
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
        this._notifySettingsChanged(sheet);
    }

    /**
     * Set the print title row/column
     * @param {{rows?: string|Array<number>|null, cols?: string|Array<number>|null}} titles - Repeated rows and columns.
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {void}
     */
    setPrintTitles(titles, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings) return;
        for (const [key, value] of Object.entries(titles || {})) {
            if (value && ((key === 'rows' && !this._normalizeTitleRowsValue(value, sheet))
                || (key === 'cols' && !this._normalizeTitleColsValue(value, sheet)))) {
                throw new Error(this.SN.t('core.print.print.invalidTitles'));
            }
        }
        settings.printTitles = this._normalizePrintTitles(titles, sheet, settings.printTitles);
        this._notifySettingsChanged(sheet);
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
        this._notifySettingsChanged(sheet);
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
        this._notifySettingsChanged(sheet);
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
        this.SN.Layout?.refreshToolbar();
    }

    /**
     * Insert Manual Page Break in the current active cell
     * @param {{r: number, c: number}|null} [cell=null] - Cell above/left of which to insert breaks.
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {{row:boolean, col:boolean}|null}
     */
    insertPageBreak(cell = null, sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings || !sheet) return null;
        const target = cell || sheet.activeCell;
        if (!target || !Number.isInteger(target.r) || !Number.isInteger(target.c)
            || target.r < 0 || target.c < 0 || target.r >= sheet.rowCount || target.c >= sheet.colCount) return null;

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

        this._notifySettingsChanged(sheet);
        return { row: rowInserted, col: colInserted };
    }

    /** @param {Sheet} [sheet=this.SN.activeSheet] Remove all manual page breaks. */
    resetPageBreaks(sheet = this.SN.activeSheet) {
        this.setSettings({ pageBreaks: { rowBreaks: [], colBreaks: [] } }, sheet);
    }

    /**
     * Get page break guide information (for canvas display)
     * @param {Sheet} [sheet=this.SN.activeSheet] - Sheet
     * @returns {{range:Object,rowBreaks:Array<{index:number,manual:boolean}>,colBreaks:Array<{index:number,manual:boolean}>}|null}
     */
    getPageBreakGuides(sheet = this.SN.activeSheet) {
        const settings = this.ensureSettings(sheet);
        if (!settings || !sheet) return null;
        let layout;
        try {
            layout = this.buildLayout(sheet, settings);
        } catch {
            return null;
        }
        const range = layout.range;
        if (!range) return null;

        const fitToPage = settings.pageSetup.fitToWidth > 0 || settings.pageSetup.fitToHeight > 0;
        const manualRowSet = new Set(
            (fitToPage ? [] : settings.pageBreaks?.rowBreaks || [])
                .map(val => toInt(val, null))
                .filter(val => Number.isFinite(val))
                .filter(val => val > range.s.r && val <= range.e.r)
        );
        const manualColSet = new Set(
            (fitToPage ? [] : settings.pageBreaks?.colBreaks || [])
                .map(val => toInt(val, null))
                .filter(val => Number.isFinite(val))
                .filter(val => val > range.s.c && val <= range.e.c)
        );

        // Show only effective page boundaries; breaks inside repeated titles do not split pages.
        const rowBreakSet = new Set();
        const colBreakSet = new Set();
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
     * Open the print settings and preview panel.
     * @param {Object} [options={}] - Preview Options
     * @returns {Promise<void>}
     */
    async preview(options = {}) {
        if (this._dialog || !(options.sheet || this.SN.activeSheet)) return;
        const dialog = new PrintDialog(this.SN, options);
        this._dialog = dialog;
        try {
            await dialog.open();
        } finally {
            if (this._dialog === dialog) this._dialog = null;
        }
    }

    /**
     * Print a job through the browser's system print dialog.
     * @param {Object} [options={}] - Job options, or {job, fromPage, toPage, signal, onProgress}.
     * @returns {Promise<void>}
     */
    async print(options = {}) {
        const job = options.job || this.createJob(options);
        if (!job.pages.length) {
            this.SN.Utils.toast(this.SN.t('core.print.print.toast.nothingToPrint'));
            return;
        }
        const fromPage = Number(options.fromPage ?? 1);
        const toPage = Number(options.toPage ?? job.pages.length);
        if (!Number.isInteger(fromPage) || !Number.isInteger(toPage)
            || fromPage < 1 || toPage < fromPage || toPage > job.pages.length) {
            throw new Error(this.SN.t('core.print.print.invalidPageRange', { total: job.pages.length }));
        }
        const selected = job.pages.slice(fromPage - 1, toPage);
        const pages = [];
        for (const page of selected) {
            options.signal?.throwIfAborted();
            pages.push(await this.renderPage(page, { signal: options.signal }));
            options.onProgress?.(pages.length, selected.length);
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        options.signal?.throwIfAborted();
        await this._printDocument(this.buildPrintHtml(pages), options.signal);
    }

    /**
     * Create a pagination snapshot without rendering images or changing saved print areas.
     * @param {Object} [options={}] - {sheet, scope: 'sheet'|'selection'|'workbook', ranges, ignorePrintArea}.
     * @returns {{pages: Array<Object>}}
     */
    createJob(options = {}) {
        const now = new Date();
        const dateText = now.toLocaleDateString();
        const timeText = now.toLocaleTimeString();
        const currentSheet = options.sheet || this.SN.activeSheet;
        const sheets = options.scope === 'workbook' ? this.SN.sheets.filter(sheet => !sheet.hidden) : [currentSheet];
        const pages = [];
        for (const sheet of sheets) {
            if (!sheet) continue;
            sheet._init();
            const settings = this.getSettings(sheet);
            const selection = options.ranges || (sheet.activeAreas.length ? sheet.activeAreas : [{ s: sheet.activeCell, e: sheet.activeCell }]);
            const ranges = options.scope === 'selection'
                ? selection.map(range => this._selectionRange(range, sheet)).filter(Boolean)
                : [null];
            for (const range of ranges) {
                const pageSettings = clone(settings);
                if (range) {
                    pageSettings.printArea = range;
                    pageSettings.printTitles = { rows: null, cols: null };
                } else if (options.ignorePrintArea) pageSettings.printArea = null;
                const layout = this.buildLayout(sheet, pageSettings);
                pages.push(...this._describePages(sheet, layout));
            }
        }
        pages.forEach((page, index) => {
            page.pageNumber = index + 1;
            page.totalPages = pages.length;
            page.dateText = dateText;
            page.timeText = timeText;
        });
        return { pages };
    }

    /**
     * Build page break layout information
     * @param {Sheet} sheet - Sheet
     * @param {Object} settings - Print Settings
     * @returns {Object}
     */
    buildLayout(sheet, settings) {
        const paper = this.getPaperDefinition(settings);
        const margins = this.getMarginPixels(settings.pageMargins);
        const printable = {
            width: Math.max(0, paper.widthPx - margins.left - margins.right),
            height: Math.max(0, paper.heightPx - margins.top - margins.bottom)
        };
        const range = this.getPrintRange(sheet, settings);
        if (!range) return { paper, margins, printable, range: null, rowSegments: [], colSegments: [], settings };
        if (printable.width <= 0 || printable.height <= 0) {
            throw new Error(this.SN.t('core.print.print.invalidMargins'));
        }
        const scale = this._resolvePrintScale(sheet, settings, printable);
        const headings = settings.printOptions?.headings;
        const contentLimit = {
            width: printable.width / scale - (headings ? sheet._indexWidth : 0),
            height: printable.height / scale - (headings ? sheet._headHeight : 0)
        };
        const repeatTitles = this._resolveRepeatTitles(sheet, settings, range);
        const bodyRange = this._buildBodyRange(range, repeatTitles);
        const repeatWidth = repeatTitles.cols && (repeatTitles.cols.s < range.s.c || repeatTitles.cols.e < range.e.c) ? repeatTitles.colsSize : 0;
        const repeatHeight = repeatTitles.rows && (repeatTitles.rows.s < range.s.r || repeatTitles.rows.e < range.e.r) ? repeatTitles.rowsSize : 0;
        if (repeatWidth >= contentLimit.width || repeatHeight >= contentLimit.height) {
            throw new Error(this.SN.t('core.print.print.titlesTooLarge'));
        }
        const bodyLimit = {
            width: Math.max(1, contentLimit.width - repeatWidth),
            height: Math.max(1, contentLimit.height - repeatHeight)
        };

        const fitToPage = settings.pageSetup.fitToWidth > 0 || settings.pageSetup.fitToHeight > 0;
        const rowSegments = this.buildSegments(
            bodyRange.s.r,
            bodyRange.e.r,
            row => this.getRowSize(sheet, row),
            bodyLimit.height,
            fitToPage ? [] : settings.pageBreaks?.rowBreaks || []
        );
        const colSegments = this.buildSegments(
            bodyRange.s.c,
            bodyRange.e.c,
            col => this.getColSize(sheet, col),
            bodyLimit.width,
            fitToPage ? [] : settings.pageBreaks?.colBreaks || []
        );

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
            sheetName: sheet.name,
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
            widthPx: widthIn * PRINT_DPI,
            heightPx: heightIn * PRINT_DPI
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
     * @returns {{s: {r: number, c: number}, e: {r: number, c: number}}|null}
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
        const parsed = typeof range === 'string' ? sheet.rangeStrToNum(range) : range;
        if (parsed?.s && parsed?.e) {
            const normalized = this._selectionRange(parsed, sheet);
            if (Object.values(normalized.s).concat(Object.values(normalized.e)).every(Number.isInteger)) return normalized;
        }
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

        for (const [rowIndex, row] of sheet._pendingXlsxRowMap || []) {
            const cells = Array.isArray(row.c) ? row.c : row.c ? [row.c] : [];
            for (const cell of cells) {
                if (cell.v === undefined && cell.f === undefined && cell.is === undefined && !Number(cell._$s)) continue;
                const address = this.SN.Utils.cellStrToNum(cell._$r);
                if (!Number.isFinite(address?.c)) continue;
                minR = Math.min(minR, rowIndex);
                minC = Math.min(minC, address.c);
                maxR = Math.max(maxR, rowIndex);
                maxC = Math.max(maxC, address.c);
            }
        }

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

        for (const drawing of sheet.Drawing?.getAll() || []) {
            if (!drawing.area) continue;
            minR = Math.min(minR, drawing.area.s.r);
            minC = Math.min(minC, drawing.area.s.c);
            maxR = Math.max(maxR, drawing.area.e.r);
            maxC = Math.max(maxC, drawing.area.e.c);
        }
        for (const merge of sheet.merges || []) {
            if (merge.s.r > maxR || merge.s.c > maxC || merge.e.r < minR || merge.e.c < minC) continue;
            minR = Math.min(minR, merge.s.r);
            minC = Math.min(minC, merge.s.c);
            maxR = Math.max(maxR, merge.e.r);
            maxC = Math.max(maxC, merge.e.c);
        }

        if (maxR < 0 || maxC < 0) {
            return null;
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
            if (size + nextSize > maxSegmentSize + 0.001 && i > segStart && size > 0) {
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
        if (row) return row.height;
        const state = sheet._getPendingXlsxRowState(r);
        if (state) return state.hidden || state.filterHidden ? 0 : state.height;
        const pending = sheet._pendingXlsxRowMap?.get(r);
        if (pending?._$hidden === '1') return 0;
        return pending?._$ht ? Math.round(Number(pending._$ht) * 1.333) : sheet.defaultRowHeight;
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
        const descriptors = this._describePages(sheet, layout);
        for (const page of descriptors) {
            pages.push(await this.renderPage({ ...page, pageNumber: pages.length + 1, totalPages: descriptors.length }));
        }
        return pages;
    }

    /**
     * Render one job page, including repeated titles.
     * @param {Object} page - A page descriptor from createJob.
     * @param {{signal?: AbortSignal}} [options={}] - Optional cancellation signal.
     * @returns {Promise<Object>}
     */
    async renderPage(page, options = {}) {
        options.signal?.throwIfAborted();
        const { sheet, layout, rowBands, colBands } = page;
        const headings = layout.settings.printOptions?.headings;
        const headingWidth = headings ? sheet._indexWidth : 0;
        const headingHeight = headings ? sheet._headHeight : 0;
        const outputScale = Math.min(2, 2 * layout.scale, 16384 / Math.max(page.width, page.height), Math.sqrt(16777216 / (page.width * page.height)));
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.ceil(page.width * outputScale));
        canvas.height = Math.max(1, Math.ceil(page.height * outputScale));
        const context = canvas.getContext('2d');
        let offsetY = 0;
        for (let rowIndex = 0; rowIndex < rowBands.length; rowIndex++) {
            let offsetX = 0;
            let bandHeight = 0;
            for (let colIndex = 0; colIndex < colBands.length; colIndex++) {
                options.signal?.throwIfAborted();
                const range = {
                    s: { r: rowBands[rowIndex].s, c: colBands[colIndex].s },
                    e: { r: rowBands[rowIndex].e, c: colBands[colIndex].e }
                };
                const size = this.getRangeSize(sheet, range, headings);
                const cropX = colIndex ? headingWidth : 0;
                const cropY = rowIndex ? headingHeight : 0;
                const width = size.width - cropX;
                bandHeight = size.height - cropY;
                if (width <= 0 || bandHeight <= 0) continue;
                const dataUrl = await this.captureRange(sheet, range, size.width, size.height, layout.settings, outputScale);
                const image = new Image();
                image.src = dataUrl;
                await image.decode();
                options.signal?.throwIfAborted();
                const sourceX = cropX / size.width * image.width;
                const sourceY = cropY / size.height * image.height;
                context.drawImage(image, sourceX, sourceY, image.width - sourceX, image.height - sourceY,
                    offsetX * outputScale, offsetY * outputScale, width * outputScale, bandHeight * outputScale);
                offsetX += width;
            }
            offsetY += bandHeight;
        }
        return { ...page, dataUrl: canvas.toDataURL('image/png') };
    }

    /**
     * Caption range specified
     * @param {Sheet} sheet - Sheet
     * @param {s: {r: number, c: number}, e: {r: number, c: number}}range - range
     * @param {number} width - Width
     * @param {number} height - Height
     * @param {Object} settings - Print Settings
     * @param {number} [outputScale=2] - Raster resolution multiplier.
     * @returns {Promise<string>}
     */
    async captureRange(sheet, range, width, height, settings, outputScale = 2) {
        return this.SN.Canvas.captureScreenshot(this.SN.Utils.cellNumToStr(range.s), 'topleft', {
            sheet,
            width: Math.max(1, width),
            height: Math.max(1, height),
            exportScale: outputScale,
            printOptions: settings.printOptions,
            compress: false,
            hideSelectionFrame: true
        });
    }

    /** @param {Object} page - Rendered job page. @returns {string} Paper HTML shared by preview and print. */
    buildPageHtml(page) {
        const { paper, margins, scale, settings } = page.layout;
        const headerFooter = this._buildPageHeaderFooterHtml(settings, margins, paper, {
            pageNumber: page.pageNumber,
            totalPages: page.totalPages,
            sheetName: page.layout.sheetName,
            dateText: page.dateText,
            timeText: page.timeText
        });
        const pageStyle = `width:${paper.widthPx}px;height:${paper.heightPx}px;padding:${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px;${this._buildPageStyle(settings, false)}`;
        return `<div class="sn-print-page" style="${escapeAttr(pageStyle)}">
            ${headerFooter.header}
            <img class="sn-print-page-image" src="${escapeAttr(page.dataUrl)}" style="width:${page.width * scale}px;height:${page.height * scale}px;" alt="">
            ${headerFooter.footer}
        </div>`;
    }

    /** @param {Array<Object>} pages - Rendered job pages. @returns {string} Standalone printable document. */
    buildPrintHtml(pages) {
        const sizes = new Map();
        const body = pages.map(page => {
            const paper = page.layout.paper;
            const size = `${paper.widthIn}in ${paper.heightIn}in`;
            if (!sizes.has(size)) sizes.set(size, `snPaper${sizes.size}`);
            return `<section style="page:${sizes.get(size)}">${this.buildPageHtml(page)}</section>`;
        }).join('');
        const pageRules = [...sizes].map(([size, name]) => `@page ${name} { size:${size}; margin:0; }`).join('\n');
        return `<!DOCTYPE html><html><head><meta charset="UTF-8">
            <title>${escapeHtml(this.SN.workbookName || '')}</title>
            <style>
                @page { size:${sizes.keys().next().value || 'A4'}; margin:0; }
                ${pageRules}
                body { margin:0; padding:0; }
                section { break-after:page; }
                section:last-child { break-after:auto; }
                .sn-print-page { position:relative; box-sizing:border-box; overflow:hidden; background-color:white; print-color-adjust:exact; -webkit-print-color-adjust:exact; }
                .sn-print-page-image { display:block; }
                .sn-print-page-header,.sn-print-page-footer { position:absolute; color:#6b7280; font:12px/1.4 Arial,sans-serif; white-space:pre-wrap; }
            </style></head><body>${body}</body></html>`;
    }

    _normalizePageBreaks(values, limit) {
        return [...new Set(normalizeArray(values).map(Number)
            .filter(index => Number.isInteger(index) && index > 0 && index < limit))].sort((a, b) => a - b);
    }

    _notifySettingsChanged(sheet) {
        this._dialog?._settingsChanged(sheet);
        if (this._settingsRefreshPending) return;
        this._settingsRefreshPending = true;
        queueMicrotask(() => {
            this._settingsRefreshPending = false;
            this.SN.Layout?.refreshToolbar();
            if (this.SN.activeSheet?.showPageBreaks) this.SN._r();
        });
    }

    _selectionRange(range, sheet) {
        if (!range?.s || !range?.e) return null;
        const used = [range.s.r, range.s.c, range.e.r, range.e.c].every(Number.isFinite) ? null : this.getUsedRange(sheet);
        const startRow = Math.max(0, Math.min(range.s.r, range.e.r));
        const startCol = Math.max(0, Math.min(range.s.c, range.e.c));
        const endRow = Math.max(range.s.r, range.e.r);
        const endCol = Math.max(range.s.c, range.e.c);
        return {
            s: { r: startRow, c: startCol },
            e: {
                r: Number.isFinite(endRow) ? endRow : Math.max(startRow, used?.e.r ?? 0),
                c: Number.isFinite(endCol) ? endCol : Math.max(startCol, used?.e.c ?? 0)
            }
        };
    }

    _describePages(sheet, layout) {
        if (!layout.range) return [];
        const pages = [];
        const makeBands = (segment, repeat) => {
            if (!repeat) return [segment];
            if (segment.s <= repeat.e) return [{ s: Math.min(segment.s, repeat.s), e: segment.e }];
            return [repeat, segment];
        };
        for (const rowSegment of layout.rowSegments) {
            for (const colSegment of layout.colSegments) {
                const rowBands = makeBands(rowSegment, layout.repeatTitles.rows);
                const colBands = makeBands(colSegment, layout.repeatTitles.cols);
                let width = layout.settings.printOptions?.headings ? sheet._indexWidth : 0;
                let height = layout.settings.printOptions?.headings ? sheet._headHeight : 0;
                for (const band of rowBands) {
                    for (let row = band.s; row <= band.e; row++) height += this.getRowSize(sheet, row);
                }
                for (const band of colBands) {
                    for (let col = band.s; col <= band.e; col++) width += this.getColSize(sheet, col);
                }
                if (width <= 0 || height <= 0) continue;
                pages.push({ sheet, layout, rowBands, colBands, width, height });
            }
        }
        return pages;
    }

    async _printDocument(html, signal) {
        signal?.throwIfAborted();
        const iframe = document.createElement('iframe');
        iframe.setAttribute('aria-hidden', 'true');
        iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
        await new Promise((resolve, reject) => {
            let timer;
            let finished = false;
            const finish = error => {
                if (finished) return;
                finished = true;
                clearTimeout(timer);
                signal?.removeEventListener('abort', abort);
                iframe.remove();
                if (error) reject(error);
                else resolve();
            };
            const abort = () => finish(signal.reason || new DOMException('Aborted', 'AbortError'));
            signal?.addEventListener('abort', abort, { once: true });
            iframe.onload = async () => {
                try {
                    const printWindow = iframe.contentWindow;
                    if (!printWindow) throw new Error(this.SN.t('core.print.print.failed'));
                    await Promise.all(Array.from(printWindow.document.images, image => image.decode()));
                    await printWindow.document.fonts?.ready;
                    signal?.throwIfAborted();
                    if (finished) return;
                    clearTimeout(timer);
                    printWindow.addEventListener('afterprint', () => finish(), { once: true });
                    printWindow.focus();
                    printWindow.print();
                } catch (error) {
                    finish(error);
                }
            };
            timer = setTimeout(() => finish(new Error(this.SN.t('core.print.print.failed'))), 60000);
            iframe.srcdoc = html;
            this.SN.containerDom.appendChild(iframe);
        });
    }

    _resolvePrintScale(sheet, settings, printable) {
        const setup = settings.pageSetup || {};
        const fitToWidth = Math.max(0, toInt(setup.fitToWidth, 0));
        const fitToHeight = Math.max(0, toInt(setup.fitToHeight, 0));
        if (!fitToWidth && !fitToHeight) return clampScale(setup.scale) / 100;
        const range = this.getPrintRange(sheet, settings);
        const repeats = this._resolveRepeatTitles(sheet, settings, range);
        const body = this._buildBodyRange(range, repeats);
        const repeatWidth = repeats.cols && (repeats.cols.s < range.s.c || repeats.cols.e < range.e.c) ? repeats.colsSize : 0;
        const repeatHeight = repeats.rows && (repeats.rows.s < range.s.r || repeats.rows.e < range.e.r) ? repeats.rowsSize : 0;
        const headings = settings.printOptions?.headings;
        const widthScale = fitToWidth ? this._fitAxisScale(body.s.c, body.e.c, index => this.getColSize(sheet, index),
            printable.width, repeatWidth + (headings ? sheet._indexWidth : 0), fitToWidth) : 1;
        const heightScale = fitToHeight ? this._fitAxisScale(body.s.r, body.e.r, index => this.getRowSize(sheet, index),
            printable.height, repeatHeight + (headings ? sheet._headHeight : 0), fitToHeight) : 1;
        return Math.min(widthScale, heightScale);
    }

    /**
     * Find a fitting scale while keeping whole rows or columns on each page.
     * @param {number} start - First row or column index.
     * @param {number} end - Last row or column index.
     * @param {Function} sizeGetter - Unscaled row height or column width.
     * @param {number} available - Printable size in pixels.
     * @param {number} overhead - Repeated titles and headings in pixels.
     * @param {number} targetPages - Maximum page count on this axis.
     * @returns {number}
     */
    _fitAxisScale(start, end, sizeGetter, available, overhead, targetPages) {
        const sizes = [];
        let total = 0;
        for (let index = start; index <= end; index++) {
            const size = Math.max(0, sizeGetter(index));
            if (size) sizes.push(size);
            total += size;
        }
        if (targetPages === 1) return Math.min(1, available / Math.max(1, total + overhead));
        const fits = scale => {
            const limit = available / scale - overhead;
            let pages = 1;
            let used = 0;
            for (const size of sizes) {
                if (size > limit + 0.001) return false;
                if (used + size > limit + 0.001) {
                    pages++;
                    used = 0;
                    if (pages > targetPages) return false;
                }
                used += size;
            }
            return true;
        };
        if (fits(1)) return 1;
        let lower = Math.min(1, available / Math.max(1, total + overhead));
        let upper = 1;
        for (let iteration = 0; iteration < 24; iteration++) {
            const candidate = (lower + upper) / 2;
            if (fits(candidate)) lower = candidate;
            else upper = candidate;
        }
        return lower;
    }

    _resolveRepeatTitles(sheet, settings, range) {
        const rows = this._parseTitleRows(settings.printTitles?.rows);
        const cols = this._parseTitleCols(settings.printTitles?.cols);
        let repeatRows = null;
        let repeatCols = null;

        if (rows) {
            const next = {
                s: rows.s,
                e: Math.min(range.e.r, rows.e)
            };
            if (next.s <= next.e && next.s <= range.s.r) repeatRows = next;
        }
        if (cols) {
            const next = {
                s: cols.s,
                e: Math.min(range.e.c, cols.e)
            };
            if (next.s <= next.e && next.s <= range.s.c) repeatCols = next;
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
        const startRow = repeatTitles.rows && repeatTitles.rows.e < range.e.r ? Math.max(range.s.r, repeatTitles.rows.e + 1) : range.s.r;
        const startCol = repeatTitles.cols && repeatTitles.cols.e < range.e.c ? Math.max(range.s.c, repeatTitles.cols.e + 1) : range.s.c;
        return {
            s: { r: startRow, c: startCol },
            e: { r: range.e.r, c: range.e.c }
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
        const s = toInt(sText, 0);
        const e = toInt(eText, 0);
        if (s < 1 || e < 1 || s > 1048576 || e > 1048576) return null;
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
        if (s > 16383 || e > 16383) return null;
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
        const cssUrl = trimmed.replace(/["\\]/g, '\\$&').replace(/[\r\n\f]/g, '');
        const safeUrl = forAttr ? escapeAttr(cssUrl) : cssUrl;
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

}
