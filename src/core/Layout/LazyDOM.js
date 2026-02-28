import getSvg from '../../assets/mainSvgs.js';

export default class LazyDOM {
    constructor(SN) {
        this.SN = SN;
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
            const el = document.createElement('div');
            el.className = 'rc-box sn-col-rc';
            const html = `
                <div onclick="${ns}.Action.addCols(event,'left')">Insert <input type="number" min="1" max="10" value="1"> column(s) left</div>
                <div onclick="${ns}.Action.addCols(event,'right')">Insert <input type="number" min="1" max="10" value="1"> column(s) right</div>
                <div onclick="${ns}.Action.hidCols()">Hide Columns</div>
                <div onclick="${ns}.Action.delCols()">Delete Columns</div>
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'row-context-menu': function () {
            const ns = this.ns;
            const el = document.createElement('div');
            el.className = 'rc-box sn-row-rc';
            const html = `
                <div onclick="${ns}.Action.addRows(event,'top')">Insert <input type="number" min="1" max="10" value="1"> row(s) above</div>
                <div onclick="${ns}.Action.addRows(event,'bottom')">Insert <input type="number" min="1" max="10" value="1"> row(s) below</div>
                <div onclick="${ns}.Action.hidRows()">Hide Rows</div>
                <div onclick="${ns}.Action.delRows()">Delete Rows</div>
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'cell-context-menu': function () {
            const ns = this.ns;
            const el = document.createElement('div');
            el.className = 'rc-box sn-cell-rc';
            el.onclick = () => { el.style.display = 'none'; };
            const html = `
                <div onclick="${ns}.Action.fontInversion('bold');">${getSvg('bold')} Bold</div>
                <div onclick="${ns}.Action.fontInversion('italic');">${getSvg('italic')} Italic</div>
                <div onclick="${ns}.Action.fontInversion('underline','single');">${getSvg('underline')} Underline</div>
                <div onclick="${ns}.Action.fontInversion('strike');">${getSvg('strikethrough')} Strikethrough</div>
                <div onclick="${ns}.AI.chatInput(areaDom.innerText)">${getSvg('gengduo')} More</div>
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        },

        'drawing-context-menu': function () {
            const ns = this.ns;
            const el = document.createElement('div');
            el.className = 'rc-box sn-drawing-rc';
            el.onclick = () => { el.style.display = 'none'; };
            const html = `
                <div onclick="${ns}.AI.chatInput('Please adjust drawing ID ' + ${ns}.Canvas.activeDrawing?.id + ': ')">${getSvg('monifenxi')} AI Adjust</div>
                <div onclick="${ns}.Canvas.activeDrawing.updIndex('top');${ns}._r('s');">${getSvg('alignTop')} Bring to Front</div>
                <div onclick="${ns}.Canvas.activeDrawing.updIndex('bottom');${ns}._r('s');">${getSvg('alignBottom')} Send to Back</div>
                <div onclick="${ns}.Canvas.activeDrawing.updIndex('up');${ns}._r('s');">${getSvg('plus')} Bring Forward</div>
                <div onclick="${ns}.Canvas.activeDrawing.updIndex('down');${ns}._r('s');">${getSvg('minus')} Send Backward</div>
                <div onclick="${ns}.activeSheet.Drawing.remove(${ns}.Canvas.activeDrawing.id);${ns}._r('s');">${getSvg('clear')} Delete Drawing</div>
            `;
            el.innerHTML = html;
            this.SN.containerDom.querySelector('.sn-top-layer').appendChild(el);
            return el;
        }
    };

    get dragLine() { return this.get('drag-line'); }
    get scrollTip() { return this.get('scroll-tip'); }
    get paddingType() { return this.get('padding-type'); }
    get hyperlinkBtn() { return this.get('hyperlink-btn'); }
    get hyperlinkTip() { return this.get('hyperlink-tip'); }
    get validTip() { return this.get('valid-tip'); }
    get validSelect() { return this.get('valid-select'); }
    get colContextMenu() { return this.get('col-context-menu'); }
    get rowContextMenu() { return this.get('row-context-menu'); }
    get cellContextMenu() { return this.get('cell-context-menu'); }
    get drawingContextMenu() { return this.get('drawing-context-menu'); }
}
