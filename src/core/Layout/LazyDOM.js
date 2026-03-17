import getSvg from '../../assets/mainSvgs.js';

export default class LazyDOM {
    /** @param {import('../Workbook/Workbook.js').default} SN */
    constructor(SN) {
        /** @type {import('../Workbook/Workbook.js').default} */
        this.SN = SN;
        /** @type {string} */
        this.ns = SN.namespace;
        this.cache = new Map();
    }

    /** @param {string} key @returns {HTMLElement|null} */
    get(key) {
        if (!this.cache.has(key)) {
            const creator = this.creators[key];
            if (!creator) {
                console.warn(`LazyDOM: Unknown key "${key}"`);
                return null;
            }
            const element = creator.call(this);
            this.cache.set(key, element);
        }
        return this.cache.get(key);
    }

    /** @param {string} key */
    has(key) {
        return this.cache.has(key);
    }

    /** @param {string} key */
    destroy(key) {
        const element = this.cache.get(key);
        if (element) {
            element.remove();
            this.cache.delete(key);
        }
    }

    destroyAll() {
        this.cache.forEach((element) => element.remove());
        this.cache.clear();
    }

    _bindContextMenuAutoHide(el) {
        el.onclick = (event) => {
            if (event?.target?.tagName === 'INPUT') return;
            el.style.display = 'none';
        };
    }

    _createContextMenuDividerHTML() {
        return '<div class="sn-rc-divider"></div>';
    }

    _createContextMenuItemHTML(action, labelHTML, iconHTML = '', extraClass = '') {
        const className = ['sn-rc-item', extraClass].filter(Boolean).join(' ');
        return `
            <div class="${className}" onclick="${action}">
                <span class="sn-rc-item-icon${iconHTML ? '' : ' sn-rc-item-icon-empty'}">${iconHTML}</span>
                <span class="sn-rc-item-label">${labelHTML}</span>
            </div>
        `;
    }

    /** Sync drawing context menu item visibility. */
    syncDrawingContextMenu() {
        const menu = this.cache.get('drawing-context-menu');
        if (!menu) return;
        const downloadItem = menu.querySelector('.sn-rc-download-image');
        if (!downloadItem) return;
        const drawing = this.SN.Canvas?.activeDrawing;
        const canDownload = drawing && ['chart', 'image', 'shape'].includes(drawing.type);
        downloadItem.style.display = canDownload ? '' : 'none';
    }

    creators = {
        'drag-line': function () {
            const el = document.createElement('div');
            el.className = 'sn-drag-line';
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'scroll-tip': function () {
            const el = document.createElement('div');
            el.className = 'sn-scroll-tip';
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'padding-type': function () {
            const ns = this.ns;
            const el = document.createElement('div');
            el.className = 'sn-dropdown sn-padding-type';
            const html = `
                <div class="sn-dropdown-toggle">${getSvg('biaoge')}</div>
                <ul class="sn-dropdown-menu">
                    <li onclick="${ns}.activeSheet.paddingArea(undefined,undefined,'order');${ns}._r();">Sequence Fill</li>
                    <li onclick="${ns}.activeSheet.paddingArea(undefined,undefined,'copy');${ns}._r();">Copy Cells</li>
                    <li onclick="${ns}.activeSheet.paddingArea(undefined,undefined,'format');${ns}._r();">Fill Formatting Only</li>
                    <li onclick="${ns}.activeSheet.paddingArea(undefined,undefined,'noFormat');${ns}._r();">Fill Without Formatting</li>
                </ul>
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'hyperlink-btn': function () {
            const el = document.createElement('div');
            el.className = 'sn-hyperlink-jump-btn sn-cur';
            el.innerHTML = getSvg('chongzuo');
            el.onclick = () => this.SN.activeSheet.hyperlinkJump(el);
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'hyperlink-tip': function () {
            const el = document.createElement('div');
            el.className = 'sn-hyperlink-tip';
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'valid-tip': function () {
            const el = document.createElement('div');
            el.className = 'sn-valid-tip';
            el.innerHTML = '<div class="fw-bold"></div><span></span>';
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'valid-select': function () {
            const el = document.createElement('div');
            el.className = 'sn-dropdown sn-cur bg-white sn-valid-select';
            el.innerHTML = '<div class="sn-dropdown-toggle"></div><ul class="sn-dropdown-menu"></ul>';
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'col-context-menu': function () {
            const ns = this.ns;
            const t = (key, params = {}) => this.SN.t(key, params);
            const countInput = '<input type="number" min="1" max="10" value="1">';
            const el = document.createElement('div');
            el.className = 'rc-box sn-scrollbar sn-col-rc';
            this._bindContextMenuAutoHide(el);
            const html = `
                ${this._createContextMenuItemHTML(`${ns}.Action.addCols(event,'left')`, t('contextMenu.column.insertLeft', { countInput }), getSvg('plus'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.addCols(event,'right')`, t('contextMenu.column.insertRight', { countInput }), getSvg('plus'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.hidCols()`, t('contextMenu.column.hide'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.unhideColAction()`, t('contextMenu.column.unhide'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.delCols()`, t('contextMenu.column.delete'), getSvg('clear'), 'sn-rc-item-danger')}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.setColWidthDialog()`, t('contextMenu.column.width'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.autoFitColWidth()`, t('contextMenu.column.autoFit'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.groupRows()`, t('contextMenu.column.group'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.ungroupRows()`, t('contextMenu.column.ungroup'))}
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'row-context-menu': function () {
            const ns = this.ns;
            const t = (key, params = {}) => this.SN.t(key, params);
            const countInput = '<input type="number" min="1" max="10" value="1">';
            const el = document.createElement('div');
            el.className = 'rc-box sn-scrollbar sn-row-rc';
            this._bindContextMenuAutoHide(el);
            const html = `
                ${this._createContextMenuItemHTML(`${ns}.Action.addRows(event,'top')`, t('contextMenu.row.insertAbove', { countInput }), getSvg('plus'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.addRows(event,'bottom')`, t('contextMenu.row.insertBelow', { countInput }), getSvg('plus'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.hidRows()`, t('contextMenu.row.hide'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.unhideRowAction()`, t('contextMenu.row.unhide'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.delRows()`, t('contextMenu.row.delete'), getSvg('clear'), 'sn-rc-item-danger')}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.setRowHeightDialog()`, t('contextMenu.row.height'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.autoFitRowHeight()`, t('contextMenu.row.autoFit'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.groupRows()`, t('contextMenu.row.group'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.ungroupRows()`, t('contextMenu.row.ungroup'))}
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'cell-context-menu': function () {
            const ns = this.ns;
            const t = (key, params = {}) => this.SN.t(key, params);
            const el = document.createElement('div');
            el.className = 'rc-box sn-scrollbar sn-cell-rc';
            this._bindContextMenuAutoHide(el);
            const html = `
                ${this._createContextMenuItemHTML(`${ns}.Action.copy()`, t('contextMenu.cell.copy'), getSvg('sort_copy'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.cut()`, t('contextMenu.cell.cut'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.paste()`, t('contextMenu.cell.paste'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.activeSheet.mergeCells(null,'center');${ns}._r()`, t('contextMenu.cell.mergeCenter'))}
                ${this._createContextMenuItemHTML(`${ns}.activeSheet.mergeCells();${ns}._r()`, t('contextMenu.cell.merge'))}
                ${this._createContextMenuItemHTML(`${ns}.activeSheet.unMergeCells();${ns}._r()`, t('contextMenu.cell.unmerge'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.insertHyperlink()`, t('contextMenu.cell.hyperlink'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.insertComment()`, t('contextMenu.cell.note'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.sortAsc()`, t('contextMenu.cell.sortAsc'), getSvg('filter_sortAsc'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.sortDesc()`, t('contextMenu.cell.sortDesc'), getSvg('filter_sortDesc'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.toggleAutoFilter()`, t('contextMenu.cell.autoFilter'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.clearFilter()`, t('contextMenu.cell.clearFilter'), getSvg('clear'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.toggleWrapText()`, t('contextMenu.cell.wrapText'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.clearAreaFormat()`, t('contextMenu.cell.clearFormat'), getSvg('clear'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.fontInversion('bold');`, t('contextMenu.cell.bold'), getSvg('bold'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.fontInversion('italic');`, t('contextMenu.cell.italic'), getSvg('italic'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.fontInversion('underline','single');`, t('contextMenu.cell.underline'), getSvg('underline'))}
                ${this._createContextMenuItemHTML(`${ns}.Action.fontInversion('strike');`, t('contextMenu.cell.strikethrough'), getSvg('strikethrough'))}
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'drawing-context-menu': function () {
            const ns = this.ns;
            const t = (key, params = {}) => this.SN.t(key, params);
            const aiAdjustPrompt = JSON.stringify(t('contextMenu.drawing.aiAdjustPrompt', { id: '__SN_DRAWING_ID__' })).replace(/"/g, '&quot;');
            const el = document.createElement('div');
            el.className = 'rc-box sn-scrollbar sn-drawing-rc';
            this._bindContextMenuAutoHide(el);
            const html = `
                ${this._createContextMenuItemHTML(`${ns}.AI.chatInput(${aiAdjustPrompt}.replace('__SN_DRAWING_ID__', ${ns}.Canvas.activeDrawing?.id ?? ''))`, t('contextMenu.drawing.aiAdjust'), getSvg('monifenxi'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Canvas.activeDrawing.updIndex('top');${ns}._r('s');`, t('contextMenu.drawing.bringToFront'), getSvg('alignTop'))}
                ${this._createContextMenuItemHTML(`${ns}.Canvas.activeDrawing.updIndex('bottom');${ns}._r('s');`, t('contextMenu.drawing.sendToBack'), getSvg('alignBottom'))}
                ${this._createContextMenuItemHTML(`${ns}.Canvas.activeDrawing.updIndex('up');${ns}._r('s');`, t('contextMenu.drawing.bringForward'), getSvg('plus'))}
                ${this._createContextMenuItemHTML(`${ns}.Canvas.activeDrawing.updIndex('down');${ns}._r('s');`, t('contextMenu.drawing.sendBackward'), getSvg('minus'))}
                ${this._createContextMenuDividerHTML()}
                ${this._createContextMenuItemHTML(`${ns}.Action.downloadActiveDrawingImage()`, t('contextMenu.drawing.downloadImage'), '', 'sn-rc-download-image')}
                ${this._createContextMenuItemHTML(`${ns}.activeSheet.Drawing.remove(${ns}.Canvas.activeDrawing.id);${ns}._r('s');`, t('contextMenu.drawing.delete'), getSvg('clear'), 'sn-rc-item-danger')}
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        }
    };

    /** @type {HTMLElement|null} */
    get dragLine() { return this.get('drag-line'); }
    /** @type {HTMLElement|null} */
    get scrollTip() { return this.get('scroll-tip'); }
    /** @type {HTMLElement|null} */
    get paddingType() { return this.get('padding-type'); }
    /** @type {HTMLElement|null} */
    get hyperlinkBtn() { return this.get('hyperlink-btn'); }
    /** @type {HTMLElement|null} */
    get hyperlinkTip() { return this.get('hyperlink-tip'); }
    /** @type {HTMLElement|null} */
    get validTip() { return this.get('valid-tip'); }
    /** @type {HTMLElement|null} */
    get validSelect() { return this.get('valid-select'); }
    /** @type {HTMLElement|null} */
    get colContextMenu() { return this.get('col-context-menu'); }
    /** @type {HTMLElement|null} */
    get rowContextMenu() { return this.get('row-context-menu'); }
    /** @type {HTMLElement|null} */
    get cellContextMenu() { return this.get('cell-context-menu'); }
    /** @type {HTMLElement|null} */
    get drawingContextMenu() { return this.get('drawing-context-menu'); }
}
