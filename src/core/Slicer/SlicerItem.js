import getSvg from '../../assets/mainSvgs.js';

const DEFAULT_MIN_WIDTH = 140;
const DEFAULT_MIN_HEIGHT = 160;

/** @class */
export default class SlicerItem {
    /** @param {Sheet} sheet @param {Object} options */
    constructor(sheet, options = {}) {
        this.sheet = sheet;
        /** @type {string} */
        this.id = options.id ?? `Slicer_${Date.now()}`;
        /** @type {string} */
        this.name = options.name ?? this.id;
        /** @type {string} */
        this.caption = options.caption ?? options.fieldName ?? '';
        /** @type {string} */
        this.cacheId = options.cacheId;
        /** @type {string} */
        this.sourceType = options.sourceType ?? 'table';
        /** @type {string} */
        this.sourceSheet = options.sourceSheet ?? sheet.name;
        /** @type {string|null} */
        this.sourceName = options.sourceName ?? null;
        /** @type {string|null} */
        this.sourceId = options.sourceId ?? null;
        /** @type {number} */
        this.fieldIndex = Number.isFinite(options.fieldIndex) ? options.fieldIndex : 0;
        /** @type {string} */
        this.fieldName = options.fieldName ?? '';
        /** @type {number} */
        this.columns = Math.max(1, Number(options.columns) || 1);
        /** @type {number} */
        this.buttonHeight = Math.max(18, Number(options.buttonHeight) || 20);
        /** @type {boolean} */
        this.showHeader = options.showHeader !== false;
        /** @type {boolean} */
        this.multiSelect = !!options.multiSelect;
        /** @type {string} */
        this.styleName = options.styleName ?? 'SlicerStyleLight1';
        /** @type {Set} */
        this.selectedKeys = new Set(options.selectedKeys || []);

        // Accept both Drawing and drawing option names for compatibility.
        this.Drawing = options.Drawing || options.drawing || null;
        this._SN = sheet.SN;
        if (this.Drawing) {
            this.Drawing.Slicer = this;
        }

        this._canvas = null;
        this._dom = null;
        this._itemsEl = null;
        this._bodyEl = null;
        this._headerEl = null;
        this._multiBtn = null;
        this._clearBtn = null;
        this._deleteBtn = null;
        this._resizeHandle = null;
        this._cacheVersion = -1;
        this._position = null;
    }

    get cache() {
        return this._SN._slicerCaches?.get(this.cacheId) || null;
    }

    /** @param {string} key @param {boolean} multi */
    toggleSelection(key, multi = false) {
        const cache = this.cache;
        if (!cache) return;
        const items = cache.getItems();
        const allKeys = items.map(item => item.key);

        if (!multi) {
            this.selectedKeys.clear();
            this.selectedKeys.add(key);
        } else {
            if (this.selectedKeys.has(key)) {
                this.selectedKeys.delete(key);
            } else {
                this.selectedKeys.add(key);
            }
        }

        if (this.selectedKeys.size === allKeys.length || this.selectedKeys.size === 0) {
            this.selectedKeys.clear();
        }

        this._updateSelectionClasses();
        this.applyFilter();
    }

    /** @param {boolean} apply */
    clearSelection(apply = false) {
        this.selectedKeys.clear();
        this._updateSelectionClasses();
        if (apply) {
            this.applyFilter();
        }
    }

    applyFilter() {
        if (this.sourceType === 'pivot') {
            this._applyPivotFilter();
        } else {
            this._applyTableFilter();
        }
    }

    /** @returns {Object} */
    toJSON() {
        return {
            id: this.id,
            name: this.name,
            caption: this.caption,
            cacheId: this.cacheId,
            sourceType: this.sourceType,
            sourceSheet: this.sourceSheet,
            sourceName: this.sourceName,
            sourceId: this.sourceId,
            fieldIndex: this.fieldIndex,
            fieldName: this.fieldName,
            columns: this.columns,
            buttonHeight: this.buttonHeight,
            showHeader: this.showHeader,
            multiSelect: this.multiSelect,
            styleName: this.styleName,
            selectedKeys: Array.from(this.selectedKeys)
        };
    }

    _ensureDom(canvas) {
        if (this._dom) return;
        this._canvas = canvas;
        if (!canvas?.slicersCon) return;

        const root = document.createElement('div');
        root.className = 'sn-slicer';
        root.dataset.slicerId = this.id;
        root.innerHTML = `
            <div class="sn-slicer-header">
                <div class="sn-slicer-title"></div>
                <div class="sn-slicer-actions">
                    <button class="sn-slicer-btn" data-action="multi" title="Multi-select">
                        ${getSvg('plus')}
                    </button>
                    <button class="sn-slicer-btn" data-action="clear" title="Clear filter">
                        ${getSvg('clear')}
                    </button>
                    <button class="sn-slicer-btn" data-action="delete" title="Delete slicer">
                        ${getSvg('yichu')}
                    </button>
                </div>
            </div>
            <div class="sn-slicer-body sn-scrollbar">
                <div class="sn-slicer-items"></div>
            </div>
            <div class="sn-slicer-resize"></div>
        `;

        const titleEl = root.querySelector('.sn-slicer-title');
        if (titleEl) titleEl.textContent = this.caption || this.fieldName || this.name;

        this._headerEl = root.querySelector('.sn-slicer-header');
        this._bodyEl = root.querySelector('.sn-slicer-body');
        this._itemsEl = root.querySelector('.sn-slicer-items');
        this._multiBtn = root.querySelector('[data-action="multi"]');
        this._clearBtn = root.querySelector('[data-action="clear"]');
        this._deleteBtn = root.querySelector('[data-action="delete"]');
        this._resizeHandle = root.querySelector('.sn-slicer-resize');

        this._dom = root;
        this._bindDomEvents();

        canvas.slicersCon.appendChild(root);
        this._syncHeaderState();
    }

    render(canvas, x, y, w, h, inView = true) {
        this._ensureDom(canvas);
        if (!this._dom) return;

        if (!inView) {
            this._dom.style.display = 'none';
            return;
        }

        this._dom.style.display = 'block';
        this._position = { x, y, w, h };
        this._applyPosition();

        const showHeader = this.showHeader;
        if (this._headerEl) {
            this._headerEl.style.display = showHeader ? 'flex' : 'none';
        }
        if (this._bodyEl) {
            this._bodyEl.style.top = showHeader ? `${this._headerEl?.offsetHeight || 26}px` : '0';
        }

        const cache = this.cache;
        if (cache) {
            if (cache.version !== this._cacheVersion) {
                this._renderItems(cache.getItems());
                this._cacheVersion = cache.version;
            } else {
                this._updateSelectionClasses();
            }
        }
    }

    _bindDomEvents() {
        if (!this._dom) return;

        this._dom.addEventListener('mousedown', (e) => e.stopPropagation());
        this._bodyEl?.addEventListener('wheel', (e) => e.stopPropagation(), { passive: true });

        this._itemsEl?.addEventListener('click', (e) => {
            const item = e.target.closest('.sn-slicer-item');
            if (!item) return;
            const key = item.dataset.key;
            if (!key) return;
            const multi = this.multiSelect || e.ctrlKey;
            this.toggleSelection(key, multi);
        });

        this._multiBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            this.multiSelect = !this.multiSelect;
            this._syncHeaderState();
        });

        this._clearBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            this.clearSelection(true);
        });

        this._deleteBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            this.sheet?.Slicer?.remove(this.id);
            this._canvas?.r('s');
        });

        this._headerEl?.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            this._startDrag(e);
        });

        this._resizeHandle?.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            this._startResize(e);
        });
    }

    _syncHeaderState() {
        if (!this._multiBtn) return;
        this._multiBtn.classList.toggle('sn-active', this.multiSelect);
    }

    _startDrag(event) {
        if (!this._canvas || !this._position) return;
        event.preventDefault();

        const canvas = this._canvas;
        const rect = canvas.handleLayer.getBoundingClientRect();
        const zoom = canvas.activeSheet?.zoom ?? 1;
        const startX = (event.clientX - rect.left) / zoom;
        const startY = (event.clientY - rect.top) / zoom;
        const { x, y, w, h } = this._position;
        const offsetX = startX - x;
        const offsetY = startY - y;
        const sheet = this.sheet;

        const onMove = (moveEvent) => {
            const mx = (moveEvent.clientX - rect.left) / zoom;
            const my = (moveEvent.clientY - rect.top) / zoom;
            const nextX = Math.max(sheet.indexWidth, mx - offsetX);
            const nextY = Math.max(sheet.headHeight, my - offsetY);
            this._position = { x: nextX, y: nextY, w, h };
            if (this.Drawing) {
                this.Drawing.moving = true;
                this.Drawing.position = { x: nextX, y: nextY, w, h };
            }
            this._applyPosition();
        };

        const onUp = () => {
            window.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
            if (this.Drawing) this.Drawing.moving = false;
            this._syncDrawingFromPosition();
        };

        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp, { once: true });
    }

    _startResize(event) {
        if (!this._canvas || !this._position) return;
        event.preventDefault();

        const canvas = this._canvas;
        const rect = canvas.handleLayer.getBoundingClientRect();
        const { x, y, w, h } = this._position;
        const zoom = canvas.activeSheet?.zoom ?? 1;
        const startX = event.clientX;
        const startY = event.clientY;

        const onMove = (moveEvent) => {
            const deltaX = (moveEvent.clientX - startX) / zoom;
            const deltaY = (moveEvent.clientY - startY) / zoom;
            const nextW = Math.max(DEFAULT_MIN_WIDTH, w + deltaX);
            const nextH = Math.max(DEFAULT_MIN_HEIGHT, h + deltaY);
            this._position = { x, y, w: nextW, h: nextH };
            if (this.Drawing) {
                this.Drawing.moving = true;
                this.Drawing.position = { x, y, w: nextW, h: nextH };
            }
            this._applyPosition();
        };

        const onUp = () => {
            window.removeEventListener('mousemove', onMove);
            window.removeEventListener('mouseup', onUp);
            if (this.Drawing) this.Drawing.moving = false;
            this._syncDrawingFromPosition(rect);
        };

        window.addEventListener('mousemove', onMove);
        window.addEventListener('mouseup', onUp, { once: true });
    }

    _syncDrawingFromPosition() {
        if (!this._canvas || !this.Drawing || !this._position) return;
        const { x, y, w, h } = this._position;
        const area = this._canvas.getOffset(x, y, w, h);
        if (!area?.s) return;
        this.Drawing.startCell = { r: area.s.r, c: area.s.c };
        this.Drawing.offsetX = area.s.offsetX;
        this.Drawing.offsetY = area.s.offsetY;
        this.Drawing.width = w;
        this.Drawing.height = h;
        this.Drawing.position = { x, y, w, h };
        this._canvas.r('s');
    }

    _applyPosition() {
        if (!this._dom || !this._position) return;
        const { x, y, w, h } = this._position;
        // Position relative to grid viewport container.
        const offsetX = this.sheet?.indexWidth ?? 0;
        const offsetY = this.sheet?.headHeight ?? 0;
        const zoom = this._canvas?.activeSheet?.zoom ?? 1;
        this._dom.style.left = `${(x - offsetX) * zoom}px`;
        this._dom.style.top = `${(y - offsetY) * zoom}px`;
        this._dom.style.width = `${w * zoom}px`;
        this._dom.style.height = `${h * zoom}px`;
    }

    _renderItems(items = []) {
        if (!this._itemsEl) return;
        this._itemsEl.innerHTML = '';

        const fragment = document.createDocumentFragment();
        items.forEach(item => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'sn-slicer-item';
            button.dataset.key = item.key;
            button.textContent = item.text;
            button.style.height = `${this.buttonHeight}px`;
            fragment.appendChild(button);
        });
        this._itemsEl.appendChild(fragment);
        this._itemsEl.style.gridTemplateColumns = `repeat(${this.columns}, minmax(0, 1fr))`;
        this._updateSelectionClasses();
    }

    _updateSelectionClasses() {
        if (!this._itemsEl) return;
        const items = Array.from(this._itemsEl.querySelectorAll('.sn-slicer-item'));
        const selectedKeys = this.selectedKeys;
        const hasFilter = selectedKeys.size > 0;

        items.forEach(item => {
            const key = item.dataset.key;
            const selected = selectedKeys.has(key);
            item.classList.toggle('sn-selected', selected);
            item.classList.toggle('sn-unselected', hasFilter && !selected);
        });
    }

    _applyTableFilter() {
        const cache = this.cache;
        const table = cache?.getSourceTable();
        if (!table || !table.range) return;

        if (!table.autoFilterEnabled) {
            table.autoFilterEnabled = true;
        }

        const items = cache.getItems();
        const colIndex = table.range.s.c + cache.fieldIndex;
        const scopeId = table.sheet.AutoFilter.getTableScopeId(table.id);

        if (this.selectedKeys.size === 0 || this.selectedKeys.size === items.length) {
            table.sheet.AutoFilter.clearColumnFilter(colIndex, scopeId);
            return;
        }

        const values = [];
        this.selectedKeys.forEach(key => {
            const item = cache.getItemMap().get(key);
            if (item) values.push(item.value);
        });

        table.sheet.AutoFilter.setColumnFilter(colIndex, {
            type: 'values',
            values
        }, scopeId);
    }

    _applyPivotFilter() {
        const cache = this.cache;
        const pivotTable = cache?.getSourcePivotTable();
        if (!pivotTable) return;

        const fieldIndex = cache.fieldIndex;
        const field = pivotTable.pivotFields?.[fieldIndex];
        if (!field) return;

        const sharedItems = pivotTable.cache?.fields?.[fieldIndex]?.sharedItems?.values ?? [];

        if (this.selectedKeys.size === 0 || this.selectedKeys.size === sharedItems.length) {
            pivotTable.clearFieldFilter(fieldIndex);
            pivotTable.calculate();
            pivotTable.render();
            return;
        }

        const selectedShared = new Set();
        this.selectedKeys.forEach(key => {
            const item = cache.getItemMap().get(key);
            if (item && Number.isFinite(item.sharedIndex)) {
                selectedShared.add(item.sharedIndex);
            }
        });

        const hiddenMap = {};
        field.items.forEach((item, itemIndex) => {
            if (item.t && item.t !== 'default' && item.t !== 'blank') return;
            const sharedIndex = item.x;
            hiddenMap[itemIndex] = !selectedShared.has(sharedIndex);
        });

        pivotTable.setItemsHidden(fieldIndex, hiddenMap);
        pivotTable.calculate();
        pivotTable.render();
    }
}
