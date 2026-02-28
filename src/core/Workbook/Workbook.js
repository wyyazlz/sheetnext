import Sheet from "../Sheet/Sheet.js"
import Canvas from "../Canvas/Canvas.js"
import Utils from "../Utils/Utils.js"
import UndoRedo from "../UndoRedo/UndoRedo.js"
import Formula from "../Formula/Formula.js"
import Xml from "../Xml/Xml.js"
import Layout from "../Layout/Layout.js"
import Action from "../../action/Action.js"
import Print from "../Print/Print.js"
import AI from "../AI/AI.js"
import License from "../License/License.js"
import DependencyGraph from "../Formula/DependencyGraph.js"
import IO from "../IO/IO.js"
import EventEmitter from "../Event/EventEmitter.js"
import { Operation } from "../Event/Operation.js"
import { parsePivotTables, parseSlicerCaches } from "../OpenEdition/PivotSlicerStub.js"
import I18n from "../I18n/I18n.js"
import { registerLocale, getLocalePack, getAllLocalePacks } from "../../locales/registry.js"
import Enum from "../../enum/index.js"
import "../../style/base.css"
import "../../style/editor.css"
import "../../style/tools.css"
import "../../style/dropdown.css"
import "../../style/tooltip.css"
import "../../style/modal.css"
import "../../style/formula.css"
import "../../style/nameManager.css"
import "../../style/formulaAudit.css"
import "../../style/autoFilter.css"
import "../../style/sortDialog.css"
import "../../style/findReplace.css"
import "../../style/pivotPanel.css"
import "../../style/table.css"
import "../../style/slicer.css"
class SheetNext {
    static _instanceCounter = 0;
    static _pendingLocalesConsumed = false;

    /** @param {string} locale @param {Object} messages @returns {typeof SheetNext} */
    static registerLocale(locale, messages) {
        registerLocale(locale, messages);
        return SheetNext;
    }

    /** @param {string} locale @returns {Object|undefined} */
    static getLocale(locale) {
        return getLocalePack(locale);
    }

    /**
     * Create workbook instance.
     * @param {HTMLElement} dom - Editor container element.
     * @param {SheetNextInitOptions} [options={}] - Initialization options.
     * @param {string} [options.licenseKey] - License key.
     * @param {string} [options.locale='en-US'] - Initial locale (e.g. 'en-US', 'zh-CN').
     * @param {Object<string, Object>} [options.locales] - Extra locale packs keyed by locale code.
     * @param {function} [options.menuRight] - Callback `(defaultHTML: string) => string`. Receives the default right-menu HTML, return modified HTML.
     * @param {function} [options.menuList] - Callback `(config: Array<{key: string, labelKey: string, groups: Array, contextual?: boolean}>) => Array`. Receives the default toolbar panel config array, return modified array.
     * @param {string} [options.AI_URL] - AI relay endpoint URL.
     * @param {string} [options.AI_TOKEN] - Optional bearer token for AI relay endpoint.
     */
    constructor(dom, options = {}) {
        SheetNext._consumePendingLocales();
        if (!(dom instanceof HTMLElement)) throw new Error("Please pass in the correct DOM!")
        dom.style.boxSizing = "border-box"
        dom.style.background = "#eef0f2"

        this.containerDom = dom
        /** @type {string} */
        this.namespace = this._setupGlobalNamespace()
        dom.__sheetNextInstance = this
        /** @type {'auto'|'manual'} */
        this.calcMode = 'auto'
        /** @type {Sheet[]} */
        this.sheets = []
        /** @type {Object} */
        this.properties = {}

        this.Event = new EventEmitter()
        this.License = new License(this, options.licenseKey)
        this.Utils = new Utils(this)
        this.I18n = this._createI18n(options)
        this.Print = new Print(this)
        this.Action = new Action(this)
        this.Layout = new Layout(this, options)
        this.AI = new AI(this, options)
        this.Xml = new Xml(this)
        this.IO = new IO(this)
        this.Formula = new Formula(this)
        this.DependencyGraph = new DependencyGraph(this)
        this.UndoRedo = new UndoRedo(this)
        this.Canvas = new Canvas(this)

        this._readOnly = false
        this._activeSheet = null
        this._sharedFormula = {}
        this._workbookName = this.t('workbook.defaultName')
        this._Xml = null
        this._definedNames = {}
        this._pivotCaches = new Map()
        this._slicerCaches = new Map()
        this._opStack = []
        this._trackChangesAlways = false

        this._loadXmlObj()
    }

    /** @type {Sheet} */
    get activeSheet() {
        return this._activeSheet
    }

    set activeSheet(sheetObj) {
        const oldSheet = this._activeSheet;
        if (oldSheet === sheetObj) return;
        if (this.Event.hasListeners('beforeActiveSheetChange')) {
            const beforeEvent = this.Event.emit('beforeActiveSheetChange', { oldSheet, newSheet: sheetObj });
            if (beforeEvent.canceled) return;
        }
        sheetObj._init();
        this._activeSheet = sheetObj;
        this._updListDom();
        this.Canvas.applyZoom();
        this._updateZoomDisplay();
        this._r();
        if (this.Event.hasListeners('afterActiveSheetChange')) {
            this.Event.emit('afterActiveSheetChange', { oldSheet, newSheet: sheetObj });
        }
    }

    /** @type {string} */
    get workbookName() {
        return this._workbookName;
    }

    set workbookName(name) {
        if (typeof name !== 'string') return this.Utils.toast(this.t('workbook.errors.nameMustBeString'));
        name = name.trim();
        if (name === '') return this.Utils.toast(this.t('workbook.errors.nameCannotBeEmpty'));
        if (name.length > 255) return this.Utils.toast(this.t('workbook.errors.nameTooLong'));
        if (/[\\\/:\*\?"<>\|]/.test(name)) return this.Utils.toast(this.t('workbook.errors.nameInvalidChars'));

        const oldName = this._workbookName;
        if (oldName === name) return;
        if (this.Event.hasListeners('beforeWorkbookRename')) {
            const beforeEvent = this.Event.emit('beforeWorkbookRename', { oldName, newName: name });
            if (beforeEvent.canceled) return;
        }
        this._workbookName = name;

        const fileNameDom = this.containerDom.querySelector(".sn-file-name");
        if (fileNameDom) fileNameDom.innerHTML = name;
        if (this.Event.hasListeners('afterWorkbookRename')) {
            this.Event.emit('afterWorkbookRename', { oldName, newName: name });
        }
    }

    /** @type {boolean} */
    get readOnly() {
        return this._readOnly;
    }

    set readOnly(bol) {
        this._readOnly = !!bol;
        this.Layout?.StateSync?.notify('readOnly', this._readOnly);
    }

    /** @param {string} [sheetName] @returns {Sheet|null} */
    addSheet(sheetName) {
        if (!sheetName) {
            let index = 1;
            while (this._sheetNames.some(name => name.toLowerCase() === `sheet${index}`.toLowerCase())) index++;
            sheetName = `Sheet${index}`;
        } else {
            if (typeof sheetName != 'string' || sheetName.trim() === '') return this.Utils.toast(this.t('workbook.errors.sheetNameInvalidType'));
            if (this._sheetNames.some(name => name.toLowerCase() === sheetName.toLowerCase()))
                return this.Utils.toast(this.t('workbook.errors.sheetNameExists', { name: sheetName }));
            if (sheetName.length > 31) return this.Utils.toast(this.t('workbook.errors.sheetNameLength'));
            if (/[:\/\\\*\?\[\]]/.test(sheetName)) return this.Utils.toast(this.t('workbook.errors.sheetNameInvalidChars'));
        }
        if (this.Event.hasListeners('beforeSheetAdd')) {
            const beforeEvent = this.Event.emit('beforeSheetAdd', { name: sheetName });
            if (beforeEvent.canceled) return null;
        }
        const newSheet = new Sheet({ _$name: sheetName }, this);
        newSheet._init();
        this.sheets.push(newSheet);
        this._updListDom();
        if (this.Event.hasListeners('afterSheetAdd')) {
            this.Event.emit('afterSheetAdd', { sheet: newSheet, name: sheetName });
        }
        return newSheet
    }

    /** @param {string} name */
    delSheet(name) {
        if (this.sheets.filter(st => !st.hidden && st.name != name).length == 0) return this.Utils.toast(this.t('workbook.errors.lastVisibleSheetCannotDelete'));
        const index = this.sheets.findIndex(st => st.name == name);
        if (index === -1) return this.Utils.toast(this.t('workbook.errors.sheetNotFound', { name }));
        const sheet = this.sheets[index];
        if (this.Event.hasListeners('beforeSheetDelete')) {
            const beforeEvent = this.Event.emit('beforeSheetDelete', { sheet, name, index });
            if (beforeEvent.canceled) return;
        }
        this.sheets.splice(index, 1);
        if (this.activeSheet.name == name) this.activeSheet = this.sheets.find(s => !s.hidden);
        this._updListDom();
        if (this.Event.hasListeners('afterSheetDelete')) {
            this.Event.emit('afterSheetDelete', { sheet, name, index });
        }
    }

    /** @param {string} sheetName @returns {Sheet|null} */
    getSheet(sheetName) {
        const sheet = this.sheets.find(sheet => sheet.name == sheetName)
        if (!sheet) return null
        return sheet
    }

    /** @param {boolean} [currentSheetOnly=false] */
    recalculate(currentSheetOnly = false) {
        if (!this.DependencyGraph) return;

        if (currentSheetOnly) {
            const sheet = this.activeSheet;
            if (!sheet) return;

            for (let r = 0; r < sheet.rowCount; r++) {
                const row = sheet.rows[r];
                if (!row) continue;
                for (let c = 0; c < row.cells.length; c++) {
                    const cell = row.cells[c];
                    if (cell && cell.isFormula) {
                        cell._calcVal = undefined;
                        cell._showVal = undefined;
                        cell._fmtResult = undefined;
                    }
                }
            }

            this.DependencyGraph.rebuildForSheet(sheet);
        } else {
            this.DependencyGraph.recalculateVolatile();

            for (const sheet of this.sheets) {
                if (!sheet.initialized) continue;
                for (let r = 0; r < sheet.rowCount; r++) {
                    const row = sheet.rows[r];
                    if (!row) continue;
                    for (let c = 0; c < row.cells.length; c++) {
                        const cell = row.cells[c];
                        if (cell && cell.isFormula) {
                            cell._calcVal = undefined;
                            cell._showVal = undefined;
                            cell._fmtResult = undefined;
                            const _ = cell.calcVal;
                        }
                    }
                }
            }
        }

        this._r();
    }

    get _sheetNames() {
        return this.sheets.map(st => st.name)
    }

    get _currentOperation() {
        return this._opStack[this._opStack.length - 1] || null;
    }

    /** @param {string} locale */
    setLocale(locale) {
        this.I18n.setLocale(locale);
        this.Layout?.refreshToolbar?.();
        this._updateZoomDisplay();
    }

    /** @type {string} */
    get locale() {
        return this.I18n.locale;
    }

    /** @param {string} key @param {Record<string, any>} [params={}] @returns {string} */
    t(key, params = {}) {
        return this.I18n.t(key, params);
    }

    _createI18n(options = {}) {
        const i18n = new I18n(options.locale || 'en-US', 'en-US');
        const localePacks = getAllLocalePacks();
        for (const [locale, pack] of Object.entries(localePacks)) {
            i18n.register(locale, pack);
        }

        if (options.locales && typeof options.locales === 'object') {
            for (const [locale, pack] of Object.entries(options.locales)) {
                i18n.register(locale, pack);
            }
        }
        return i18n;
    }

    static _consumePendingLocales() {
        if (SheetNext._pendingLocalesConsumed) return;
        SheetNext._pendingLocalesConsumed = true;
        const root = typeof globalThis !== 'undefined' ? globalThis : (typeof window !== 'undefined' ? window : null);
        if (!root) return;
        const pending = root.__SHEETNEXT_PENDING_LOCALES__;
        if (!pending || typeof pending !== 'object') return;
        for (const [locale, pack] of Object.entries(pending)) {
            registerLocale(locale, pack);
        }
        delete root.__SHEETNEXT_PENDING_LOCALES__;
    }

    _setupGlobalNamespace() {
        const namespace = `SN_${SheetNext._instanceCounter++}`;
        window[namespace] = this;
        return namespace;
    }

    _loadXmlObj() {
        this.Xml.obj = new Proxy(this.Xml.obj, {
            get: (target, property) => {
                if (target[property]?._data && (property.endsWith('.xml') || property.endsWith('.rels'))) {
                    return target[property] = Xml.parser.parse(target[property].asText());
                } else {
                    return target[property];
                }
            }
        });

        this._loadProperties();
        this._loadDefinedNames();

        this.sheets = this.Xml.sheets.map(ct => new Sheet(ct, this));

        this._pivotCaches = new Map();
        parsePivotTables(this, this.Xml);
        parseSlicerCaches(this, this.Xml);

        const si = this.Xml.obj["xl/workbook.xml"].workbook?.bookViews?.workbookView?.['_$activeTab'] ?? 0
        this.activeSheet = this.sheets[si]
    }

    _loadProperties() {
        const xmlObj = this.Xml.obj;
        const core = xmlObj['docProps/core.xml']?.['cp:coreProperties'] || {};
        const app = xmlObj['docProps/app.xml']?.Properties || {};

        const parseTime = (obj) => {
            if (!obj) return '';
            if (typeof obj === 'string') return obj;
            return obj['#text'] || '';
        };

        this.properties = {
            title: core['dc:title'] || '',
            subject: core['dc:subject'] || '',
            creator: core['dc:creator'] || '',
            keywords: core['cp:keywords'] || '',
            description: core['dc:description'] || '',
            lastModifiedBy: core['cp:lastModifiedBy'] || '',
            created: parseTime(core['dcterms:created']),
            modified: parseTime(core['dcterms:modified']),
            category: core['cp:category'] || '',
            contentStatus: core['cp:contentStatus'] || '',
            revision: core['cp:revision'] || '',
            language: core['dc:language'] || '',
            lastPrinted: parseTime(core['cp:lastPrinted']),
            application: app.Application || '',
            appVersion: app.AppVersion || '',
            company: app.Company || '',
            manager: app.Manager || '',
            hyperlinkBase: app.HyperlinkBase || '',
            template: app.Template || ''
        };
    }

    _loadDefinedNames() {
        const defined = this.Xml.obj["xl/workbook.xml"].workbook?.definedNames?.definedName;
        const list = Array.isArray(defined) ? defined : (defined ? [defined] : []);
        const map = {};
        list.forEach((item) => {
            const name = item?._$name;
            const ref = item?.['#text'];
            if (name && typeof ref === 'string') {
                map[name] = ref;
            }
        });
        this._definedNames = map;
    }

    _updListDom() {
        const ns = this.namespace;
        const sheetName = this.activeSheet.name
        let sheetsHtml = ""
        this.sheets.forEach((item, index) => {
            if (item.name == sheetName || this._sheetNames.length == 1) {
                sheetsHtml += `<button class="sheet-sel sn-sheet-item" onmousedown="this.classList.add('sheet-sel')" ontouchstart="this.classList.add('sheet-sel')">${item.name}</button>`
            } else if (!item.hidden) {
                sheetsHtml += `<button class="sn-sheet-item" onmousedown="${ns}.activeSheet=${ns}.getSheet('${item.name}');this.classList.add('sheet-sel')" ontouchstart="${ns}.activeSheet=${ns}.getSheet('${item.name}');this.classList.add('sheet-sel')">${item.name}</button>`
            }
        })
        const scrollContainer = this.containerDom.querySelector('.sn-sheet-scroll');
        if (scrollContainer) {
            scrollContainer.innerHTML = sheetsHtml;
        } else {
            this.containerDom.querySelector('.sn-sheet-list').innerHTML = `<div class="sn-sheet-scroll">${sheetsHtml}</div>`;
        }
    }

    _r() {
        this.Canvas.r();
    }

    _updateZoomDisplay() {
        const zoomVal = this.containerDom.querySelector('.sn-zoom-val');
        if (!zoomVal) return;
        const zoom = this.activeSheet?.zoom ?? 1;
        const percent = Math.round(zoom * 100);
        if (zoomVal.tagName === 'INPUT') {
            if (document.activeElement === zoomVal) return;
            zoomVal.value = percent;
            return;
        }
        zoomVal.textContent = `${percent}%`;
    }

    _enableChangeTracking(enabled = true) {
        this._trackChangesAlways = !!enabled;
        return this;
    }

    _beginOperation(type, data = {}, options = {}) {
        const parent = this._currentOperation;
        const trackChanges = options.trackChanges ?? (
            this._trackChangesAlways ||
            this.Event.hasListeners('afterOperation') ||
            this.Event.hasListeners(`afterOperation:${type}`)
        );
        const op = new Operation(type, data, {
            parentId: parent?.id ?? null,
            trackChanges
        });
        const beforeEvent = this.Event.emit('beforeOperation', { operation: op, ...data });
        const beforeTyped = this.Event.emit(`beforeOperation:${type}`, { operation: op, ...data });
        const cancelable = options.cancelable !== false;
        if (cancelable && (beforeEvent.canceled || beforeTyped.canceled)) return null;
        this._opStack.push(op);
        return op;
    }

    _endOperation(op, options = {}) {
        if (!op) return null;
        const current = this._currentOperation;
        if (current === op) {
            this._opStack.pop();
        } else {
            const idx = this._opStack.lastIndexOf(op);
            if (idx !== -1) this._opStack.splice(idx, 1);
        }
        op.end();
        const payload = { operation: op, changes: op.changes, summary: op.changes?.summary };
        const afterGlobal = this.Event.emitAsync('afterOperation', payload);
        const afterTyped = this.Event.emitAsync(`afterOperation:${op.type}`, payload);
        if (options.await) {
            return Promise.all([afterGlobal, afterTyped]).then(() => op);
        }
        return op;
    }

    _withOperation(type, data, fn, options = {}) {
        const parent = this._currentOperation;
        const joinParent = options.joinParent !== false;
        if (parent && joinParent) {
            return fn(parent);
        }
        const op = this._beginOperation(type, data, options);
        if (!op) return;
        try {
            return fn(op);
        } finally {
            this._endOperation(op, options);
        }
    }

    _recordChange(change) {
        const op = this._currentOperation;
        if (!op || !op.changes) return;
        op.addChange(change);
    }

}

SheetNext._consumePendingLocales();
window.SheetNext = SheetNext
SheetNext.Enum = Enum;

export default SheetNext

console.log('%c SheetNext Editor 1.0', 'background:#36c;color:#fff;padding:5px 10px;font-size:12px;');
