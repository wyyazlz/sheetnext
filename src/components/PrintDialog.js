import getSvg from '../assets/mainSvgs.js';

function escapeHtml(value) {
    return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

export default class PrintDialog {
    constructor(SN, options = {}) {
        this._SN = SN;
        this._print = SN.Print;
        this._sheet = options.sheet || SN.activeSheet;
        this._settingsSheet = this._sheet;
        this._ranges = (options.ranges || this._sheet.activeAreas).map(range => ({ s: { ...range.s }, e: { ...range.e } }));
        if (!this._ranges.length) this._ranges.push({ s: { ...this._sheet.activeCell }, e: { ...this._sheet.activeCell } });
        this._scope = options.scope || 'sheet';
        this._body = null;
        this._job = null;
        this._pageIndex = 0;
        this._version = 0;
        this._cache = new Map();
        this._timer = null;
        this._renderPromise = null;
        this._pendingRender = false;
        this._displayPage = null;
        this._printing = false;
        this._applyingSetting = false;
        this._closed = false;
        this._events = new AbortController();
        this._printAbort = null;
        this._previewAbort = new AbortController();
        this._observer = null;
        this._previousFocus = document.activeElement;
    }

    async open() {
        this._sheet._init();
        await this._SN.Utils.modal({
            titleKey: 'core.print.print.modal.m001.title',
            content: this._buildHtml(),
            modalClass: 'sn-print-dialog',
            bodyClass: 'sn-print-workspace',
            maxWidth: '1400px',
            maxHeight: '96vh',
            onOpen: body => this._mount(body),
            onClosed: () => this._dispose()
        });
        this._dispose();
    }

    _t(key, params) {
        return this._SN.t(`core.print.print.${key}`, params);
    }

    _label(key) {
        return escapeHtml(this._t(key));
    }

    _field(name) {
        return this._body.querySelector(`[name="${name}"]`);
    }

    _select(name, label, choices, value) {
        return `<label class="sn-print-field"><span>${escapeHtml(label)}</span><select name="${name}">
            ${choices.map(([option, text]) => `<option value="${escapeHtml(option)}"${String(option) === String(value) ? ' selected' : ''}>${escapeHtml(text)}</option>`).join('')}
        </select></label>`;
    }

    _input(name, label, value, attributes = '') {
        return `<label class="sn-print-field"><span>${escapeHtml(label)}</span><input name="${name}" value="${escapeHtml(value)}" ${attributes}></label>`;
    }

    _checkbox(name, label, checked) {
        return `<label class="sn-print-check"><input type="checkbox" name="${name}"${checked ? ' checked' : ''}><span>${escapeHtml(label)}</span></label>`;
    }

    _buildHtml() {
        return `<aside class="sn-print-sidebar">
            <button type="button" class="sn-print-submit sn-modal-btn" data-command="print">${getSvg('dayinyulan')}<span>${this._label('modal.m001.confirm')}</span></button>
            <p class="sn-print-hint">${this._label('systemPrinter')}</p>
            <fieldset class="sn-print-options">
                <h3>${this._label('settings')}</h3>
                ${this._select('scope', this._t('scope'), ['sheet', 'selection', 'workbook'].map(scope => [scope, this._t(`scopeOptions.${scope}`)]), this._scope)}
                <div class="sn-print-row sn-print-page-range-inputs">
                    ${this._input('fromPage', this._t('fromPage'), '', `type="number" min="1" step="1" placeholder="${this._label('firstPage')}"`)}
                    ${this._input('toPage', this._t('toPage'), '', `type="number" min="1" step="1" placeholder="${this._label('lastPage')}"`)}
                </div>
                ${this._checkbox('ignorePrintArea', this._t('ignorePrintArea'), false)}
                <div class="sn-print-sheet-settings"></div>
            </fieldset>
            <p class="sn-print-feedback" role="status" aria-live="polite"></p>
        </aside>
        <main class="sn-print-preview-pane">
            <div class="sn-print-preview-caption"></div>
            <div class="sn-print-viewport" aria-busy="true"><div class="sn-print-message">${this._label('loading')}</div></div>
            <div class="sn-print-navigation">
                <div class="sn-print-pager">
                    <button type="button" data-command="previous" aria-label="${this._label('previousPage')}">‹</button>
                    <input name="pageNumber" type="number" min="1" step="1" value="1" aria-label="${this._label('pageNumber')}">
                    <span class="sn-print-page-count"></span>
                    <button type="button" data-command="next" aria-label="${this._label('nextPage')}">›</button>
                </div>
                ${this._select('previewZoom', this._t('previewZoom'), [['fit', this._t('fitPage')], ['0.5', '50%'], ['0.75', '75%'], ['1', '100%'], ['1.5', '150%']], 'fit')}
            </div>
        </main>`;
    }

    _buildSettings() {
        const settings = this._print.getSettings(this._settingsSheet);
        const setup = settings.pageSetup;
        const margins = settings.pageMargins;
        const presets = this._print.getMarginPresets();
        const preset = this._print.getPageMarginsPreset(this._settingsSheet);
        const scaling = setup.fitToWidth || setup.fitToHeight
            ? (setup.fitToWidth === 1 && setup.fitToHeight === 1 ? 'sheet' : setup.fitToWidth === 1 && !setup.fitToHeight ? 'columns' : !setup.fitToWidth && setup.fitToHeight === 1 ? 'rows' : 'fitCustom')
            : setup.scale === 100 ? 'none' : 'custom';
        const area = settings.printArea ? (typeof settings.printArea === 'string' ? settings.printArea : this._SN.Utils.rangeNumToStr(settings.printArea)) : this._t('usedArea');
        return `${this._scope === 'workbook' ? this._select('settingsSheet', this._t('settingsSheet'), this._SN.sheets.filter(sheet => !sheet.hidden).map(sheet => [this._SN.sheets.indexOf(sheet), sheet.name]), this._SN.sheets.indexOf(this._settingsSheet)) : ''}
            <p class="sn-print-hint">${escapeHtml(this._scope === 'selection' ? this._t('selectionOnly') : this._t('savedArea', { area }))}</p>
            ${this._select('orientation', this._SN.t('layout.page.orientation'), [['portrait', this._t('portrait')], ['landscape', this._t('landscape')]], setup.orientation)}
            ${this._select('paperName', this._SN.t('layout.page.paperSize'), this._print.getPaperSizes().map(paper => [paper.name, `${paper.name} · ${(paper.widthIn * 2.54).toFixed(1)} × ${(paper.heightIn * 2.54).toFixed(1)} ${this._t('cm')}`]), setup.paperName)}
            ${this._select('marginPreset', this._SN.t('layout.page.margins'), [...Object.keys(presets), 'custom'].map(name => [name, this._t(`margins.${name}`)]), preset)}
            <div class="sn-print-margin-fields sn-print-grid"${preset === 'custom' ? '' : ' hidden'}>
                ${['top', 'bottom', 'left', 'right', 'header', 'footer'].map(field => this._input(`margin-${field}`, this._t(`margins.${field}`), (margins[field] * 2.54).toFixed(2), 'type="number" min="0" step="0.01" required')).join('')}
                <span class="sn-print-hint">${this._label('marginUnit')}</span>
            </div>
            ${this._select('scaling', this._SN.t('layout.page.printScale'), ['none', 'sheet', 'columns', 'rows', 'custom', 'fitCustom'].map(mode => [mode, this._t(`scaling.${mode}`)]), scaling)}
            <div class="sn-print-scale-field"${scaling === 'custom' ? '' : ' hidden'}>
                ${this._input('scale', this._t('scalePercent'), setup.scale, 'type="number" min="10" max="400" step="1" required')}
            </div>
            <div class="sn-print-fit-fields sn-print-grid"${scaling === 'fitCustom' ? '' : ' hidden'}>
                ${this._input('fitToWidth', this._t('fitWidth'), setup.fitToWidth, 'type="number" min="0" max="999" step="1" required')}
                ${this._input('fitToHeight', this._t('fitHeight'), setup.fitToHeight, 'type="number" min="0" max="999" step="1" required')}
                <span class="sn-print-hint">${this._label('fitHint')}</span>
            </div>
            <details class="sn-print-advanced">
                <summary>${this._label('pageSetup')}</summary>
                ${this._checkbox('gridlines', this._SN.t('layout.page.printGridlines'), settings.printOptions.gridlines)}
                ${this._checkbox('headings', this._SN.t('layout.page.printHeadings'), settings.printOptions.headings)}
                ${this._input('titleRows', this._t('titleRows'), settings.printTitles.rows, `type="text" placeholder="1:1" pattern="[0-9]+:[0-9]+"${this._scope === 'selection' ? ' disabled' : ''}`)}
                ${this._input('titleCols', this._t('titleCols'), settings.printTitles.cols, `type="text" placeholder="A:A" pattern="[A-Za-z]+:[A-Za-z]+"${this._scope === 'selection' ? ' disabled' : ''}`)}
                ${this._input('header', this._t('header'), settings.headerFooter.header, 'type="text"')}
                ${this._input('footer', this._t('footer'), settings.headerFooter.footer, 'type="text"')}
                <p class="sn-print-hint">${this._label('headerFooterHint')}</p>
            </details>`;
    }

    _mount(body) {
        if (this._closed) return;
        this._body = body;
        this._renderSettings();
        body.addEventListener('change', event => this._change(event), { signal: this._events.signal });
        body.addEventListener('click', event => this._click(event), { signal: this._events.signal });
        body.closest('.sn-modal').addEventListener('keydown', event => {
            if (event.key === 'Escape') {
                event.preventDefault();
                body.closest('.sn-modal').querySelector('.sn-modal-close').click();
            } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'p') {
                event.preventDefault();
                this._submit();
            } else if (event.key === 'Tab') {
                const focusable = Array.from(body.closest('.sn-modal').querySelectorAll('button, input, select, summary'))
                    .filter(element => !element.disabled && element.getClientRects().length);
                const index = focusable.indexOf(document.activeElement);
                if ((event.shiftKey && index <= 0) || (!event.shiftKey && index === focusable.length - 1)) {
                    event.preventDefault();
                    focusable[event.shiftKey ? focusable.length - 1 : 0]?.focus();
                }
            }
            event.stopPropagation();
        }, { signal: this._events.signal });
        this._observer = new ResizeObserver(() => this._fitPreview());
        this._observer.observe(body.querySelector('.sn-print-viewport'));
        this._field('ignorePrintArea').disabled = this._scope === 'selection';
        this._refresh();
        this._field('scope').focus();
    }

    _renderSettings() {
        const container = this._body.querySelector('.sn-print-sheet-settings');
        const expanded = container.querySelector('details')?.open;
        container.innerHTML = this._buildSettings();
        if (expanded) container.querySelector('details').open = true;
        this._body.querySelector('.sn-print-options').disabled = this._printing;
    }

    _settingsChanged(sheet) {
        if (this._closed || !this._body || (this._scope !== 'workbook' && sheet !== this._sheet)) return;
        if (this._printing) this._printAbort?.abort();
        if (!this._applyingSetting && !this._printing) this._renderSettings();
        this._scheduleRefresh();
    }

    _change(event) {
        const field = event.target;
        const name = field.name;
        if (!name || this._printing) return;
        field.setCustomValidity?.('');
        if (!field.checkValidity()) {
            field.reportValidity();
            this._updateNavigation();
            return;
        }
        if (name === 'previewZoom') return this._fitPreview();
        if (name === 'pageNumber') {
            if (!field.value || !this._job?.pages.length) return this._updateNavigation();
            this._pageIndex = Math.min(this._job.pages.length - 1, Math.max(0, Number(field.value) - 1));
            this._renderCurrent();
            return;
        }
        if (name === 'fromPage' || name === 'toPage') {
            this._feedback('');
            return this._updateNavigation();
        }
        if (name === 'scope') {
            this._scope = field.value;
            this._settingsSheet = this._sheet;
            this._field('ignorePrintArea').disabled = this._scope === 'selection';
            this._renderSettings();
            this._scheduleRefresh();
            return;
        }
        if (name === 'ignorePrintArea') return this._scheduleRefresh();
        if (name === 'settingsSheet') {
            this._settingsSheet = this._SN.sheets[Number(field.value)];
            this._settingsSheet._init();
            this._renderSettings();
            return;
        }
        try {
            this._applyingSetting = true;
            this._applySetting(name, field);
        } catch (error) {
            this._feedback(error.message);
            field.setCustomValidity(error.message);
            field.reportValidity();
            this._updateNavigation();
        } finally {
            this._applyingSetting = false;
        }
    }

    _applySetting(name, field) {
        const sheet = this._settingsSheet;
        const number = fieldName => Number(this._field(fieldName).value);
        if (name === 'orientation' || name === 'paperName') this._print.setPageSetup({ [name]: field.value }, sheet);
        else if (name === 'marginPreset') {
            this._body.querySelector('.sn-print-margin-fields').hidden = field.value !== 'custom';
            if (field.value === 'custom') return;
            this._print.setPageMarginsPreset(field.value, sheet);
            const margins = this._print.getSettings(sheet).pageMargins;
            for (const [key, value] of Object.entries(margins)) {
                this._field(`margin-${key}`).value = (value * 2.54).toFixed(2);
                this._field(`margin-${key}`).setCustomValidity('');
            }
        } else if (name.startsWith('margin-')) {
            const inputs = Array.from(this._body.querySelectorAll('.sn-print-margin-fields input'));
            inputs.forEach(input => input.setCustomValidity(''));
            if (inputs.some(input => !input.reportValidity())) return;
            this._print.setPageMargins(Object.fromEntries(inputs.map(input => [input.name.slice(7), Number(input.value) / 2.54])), sheet);
        } else if (['scaling', 'scale', 'fitToWidth', 'fitToHeight'].includes(name)) {
            const mode = this._field('scaling').value;
            this._body.querySelector('.sn-print-scale-field').hidden = mode !== 'custom';
            this._body.querySelector('.sn-print-fit-fields').hidden = mode !== 'fitCustom';
            const setup = { scale: 100, fitToWidth: 0, fitToHeight: 0 };
            if (mode === 'sheet' || mode === 'columns') setup.fitToWidth = 1;
            if (mode === 'sheet' || mode === 'rows') setup.fitToHeight = 1;
            if (mode === 'custom') setup.scale = number('scale');
            if (mode === 'fitCustom') {
                setup.fitToWidth = number('fitToWidth');
                setup.fitToHeight = number('fitToHeight');
                if (!setup.fitToWidth && !setup.fitToHeight) {
                    setup.fitToWidth = 1;
                    this._field('fitToWidth').value = '1';
                }
            }
            this._print.setPageSetup(setup, sheet);
        } else if (name === 'gridlines' || name === 'headings') this._print.setPrintOptions({ [name]: field.checked }, sheet);
        else if (name === 'titleRows' || name === 'titleCols') {
            this._print.setPrintTitles({ [name === 'titleRows' ? 'rows' : 'cols']: field.value.trim() || null }, sheet);
        } else if (name === 'header' || name === 'footer') this._print.setSettings({ headerFooter: { [name]: field.value } }, sheet);
        this._feedback('');
        this._updateNavigation();
    }

    _click(event) {
        const command = event.target.closest('[data-command]')?.dataset.command;
        if (command === 'print') this._submit();
        if ((command === 'previous' || command === 'next') && !this._printing && !this._timer && this._job?.pages.length) {
            this._pageIndex += command === 'previous' ? -1 : 1;
            this._pageIndex = Math.max(0, Math.min(this._job.pages.length - 1, this._pageIndex));
            this._renderCurrent();
        }
    }

    _scheduleRefresh() {
        clearTimeout(this._timer);
        this._version++;
        this._previewAbort.abort();
        this._previewAbort = new AbortController();
        this._body.querySelector('[data-command="print"]').disabled = true;
        this._timer = setTimeout(() => this._refresh(), 120);
        this._updateNavigation();
    }

    _refresh() {
        if (this._closed) return;
        clearTimeout(this._timer);
        this._timer = null;
        this._version++;
        this._cache.clear();
        this._displayPage = null;
        try {
            this._job = this._print.createJob({ sheet: this._sheet, scope: this._scope, ranges: this._ranges, ignorePrintArea: this._field('ignorePrintArea').checked });
            this._pageIndex = Math.max(0, Math.min(this._pageIndex, this._job.pages.length - 1));
            for (const name of ['fromPage', 'toPage']) {
                const field = this._field(name);
                field.max = Math.max(1, this._job.pages.length);
                if (Number(field.value) > this._job.pages.length) field.value = '';
            }
            this._feedback('');
            this._renderCurrent();
        } catch (error) {
            this._job = null;
            this._message(error.message);
            this._feedback(error.message);
            this._updateNavigation();
        }
    }

    _renderCurrent() {
        this._previewAbort.abort();
        this._previewAbort = new AbortController();
        this._pendingRender = true;
        this._displayPage = null;
        this._message(this._t('loading'));
        this._updateNavigation();
        if (this._renderPromise) return;
        this._renderPromise = this._renderLoop().finally(() => { this._renderPromise = null; });
    }

    async _renderLoop() {
        while (this._pendingRender && !this._closed) {
            this._pendingRender = false;
            const page = this._job?.pages[this._pageIndex];
            const version = this._version;
            const index = this._pageIndex;
            if (!page) {
                this._message(this._t('toast.nothingToPrint'));
                this._updateNavigation();
                continue;
            }
            try {
                const rendered = this._cache.get(index) || await this._print.renderPage(page, { signal: this._previewAbort.signal });
                if (this._closed || version !== this._version || index !== this._pageIndex) continue;
                this._cache.delete(index);
                this._cache.set(index, rendered);
                if (this._cache.size > 4) this._cache.delete(this._cache.keys().next().value);
                this._displayPage = rendered;
                const viewport = this._body.querySelector('.sn-print-viewport');
                viewport.innerHTML = `<div class="sn-print-paper-frame">${this._print.buildPageHtml(rendered)}</div>`;
                viewport.setAttribute('aria-busy', 'false');
                this._fitPreview();
                const scalePercent = Math.round(page.layout.scale * 10000) / 100;
                this._body.querySelector('.sn-print-preview-caption').textContent = `${page.layout.sheetName} · ${page.layout.paper.name} · ${this._t('scalePercent')} ${scalePercent}%`;
                this._updateNavigation();
            } catch (error) {
                if (this._closed || version !== this._version || error.name === 'AbortError') continue;
                this._message(this._t('failed'));
                this._feedback(error.message || this._t('failed'));
                this._updateNavigation();
            }
        }
    }

    _fitPreview() {
        if (!this._displayPage || this._closed) return;
        const viewport = this._body.querySelector('.sn-print-viewport');
        const frame = viewport.querySelector('.sn-print-paper-frame');
        if (!frame) return;
        const paper = this._displayPage.layout.paper;
        const zoom = this._field('previewZoom').value;
        const scale = zoom === 'fit' ? Math.max(0.05, Math.min((viewport.clientWidth - 48) / paper.widthPx, (viewport.clientHeight - 48) / paper.heightPx, 1)) : Number(zoom);
        frame.style.width = `${paper.widthPx * scale}px`;
        frame.style.height = `${paper.heightPx * scale}px`;
        frame.firstElementChild.style.transform = `scale(${scale})`;
    }

    _updateNavigation() {
        if (this._closed || !this._body) return;
        const total = this._job?.pages.length || 0;
        const pending = !!this._timer || this._printing;
        this._field('pageNumber').value = total ? this._pageIndex + 1 : 0;
        this._field('pageNumber').max = Math.max(1, total);
        this._field('pageNumber').disabled = !total || pending;
        this._body.querySelector('.sn-print-page-count').textContent = this._t('pageCount', { total });
        this._body.querySelector('[data-command="previous"]').disabled = !total || pending || this._pageIndex === 0;
        this._body.querySelector('[data-command="next"]').disabled = !total || pending || this._pageIndex >= total - 1;
        const invalid = Array.from(this._body.querySelectorAll('.sn-print-options input, .sn-print-options select')).some(field => !field.closest('[hidden]') && !field.checkValidity());
        this._body.querySelector('[data-command="print"]').disabled = !this._displayPage || pending || invalid;
    }

    _message(text) {
        const viewport = this._body.querySelector('.sn-print-viewport');
        viewport.innerHTML = `<div class="sn-print-message">${escapeHtml(text)}</div>`;
        viewport.setAttribute('aria-busy', String(text === this._t('loading')));
        this._body.querySelector('.sn-print-preview-caption').textContent = '';
    }

    _feedback(text) {
        if (!this._closed) this._body.querySelector('.sn-print-feedback').textContent = text;
    }

    async _submit() {
        if (this._body?.contains(document.activeElement) && document.activeElement.matches('input')) document.activeElement.blur();
        if (this._closed || this._printing || this._timer || !this._displayPage || !this._job) return;
        const inputs = this._body.querySelectorAll('.sn-print-options input, .sn-print-options select');
        for (const input of inputs) {
            if (input.closest('[hidden]') || input.checkValidity()) continue;
            const details = input.closest('details');
            if (details) details.open = true;
            input.reportValidity();
            return;
        }
        const total = this._job.pages.length;
        const fromPage = Number(this._field('fromPage').value || 1);
        const toPage = Number(this._field('toPage').value || total);
        if (fromPage > toPage) return this._feedback(this._t('invalidPageRange', { total }));
        this._printing = true;
        this._printAbort = new AbortController();
        this._body.querySelector('.sn-print-options').disabled = true;
        this._updateNavigation();
        try {
            await this._renderPromise;
            await this._print.print({ job: this._job, fromPage, toPage, signal: this._printAbort.signal,
                onProgress: (current, count) => this._feedback(this._t('preparing', { current, total: count })) });
            this._feedback('');
        } catch (error) {
            if (error.name !== 'AbortError') this._feedback(error.message || this._t('failed'));
        } finally {
            this._printing = false;
            if (!this._closed) {
                this._body.querySelector('.sn-print-options').disabled = false;
                this._updateNavigation();
                this._body.querySelector('[data-command="print"]').focus();
            }
        }
    }

    _dispose() {
        if (this._closed) return;
        this._closed = true;
        clearTimeout(this._timer);
        this._events.abort();
        this._printAbort?.abort();
        this._previewAbort.abort();
        this._observer?.disconnect();
        this._cache.clear();
        if (this._previousFocus?.isConnected) this._previousFocus.focus?.({ preventScroll: true });
    }
}
