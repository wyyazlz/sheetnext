function parsePositiveNumber(value, fallback) {
    const num = Number(value);
    return Number.isFinite(num) && num >= 0 ? num : fallback;
}

function escapeAttr(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

export function printPreview() {
    this.SN.Print.preview();
}

export async function exportPdf() {
    await this.SN.Print.print();
}

export async function pageBreakPreview() {
    this.SN.Print.toggleShowPageBreaks(true);
    await this.SN.Print.preview({ pageBreakPreview: true });
}

export function toggleShowPageBreaks(checked) {
    this.SN.Print.toggleShowPageBreaks(!!checked);
}

export function insertPageBreak() {
    const result = this.SN.Print.insertPageBreak();
    if (!result) {
        this.SN.Utils.toast(this.SN.t('action.print.toast.pleaseSelectAValidCell'));
        return;
    }
    if (!result.row && !result.col) {
        this.SN.Utils.toast(this.SN.t('action.print.toast.aPageBreakAlreadyExists'));
        return;
    }
    this.SN.Utils.toast(this.SN.t('action.print.toast.pageBreakInserted'));
}

export function togglePrintGridlines(checked) {
    this.SN.Print.toggleGridlines(!!checked);
}

export function togglePrintHeadings(checked) {
    this.SN.Print.toggleHeadings(!!checked);
}

export function setPageMarginsPreset(preset) {
    const presets = {
        normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
        narrow: { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25, header: 0.2, footer: 0.2 },
        wide: { left: 1, right: 1, top: 1, bottom: 1, header: 0.3, footer: 0.3 }
    };
    const margin = presets[preset] || presets.normal;
    this.SN.Print.setPageMargins(margin);
    this.SN.Utils.toast(this.SN.t('action.print.toast.marginsUpdated'));
}

export async function setPageMarginsDialog() {
    const SN = this.SN;
    const settings = SN.Print.ensureSettings(SN.activeSheet);
    const margins = settings.pageMargins || {};
    const html = `
        <div class="sn-form-row">
            <label>Top</label>
            <input type="number" step="0.01" data-field="pm-top" value="${margins.top ?? 0.75}">
        </div>
        <div class="sn-form-row">
            <label>Bottom</label>
            <input type="number" step="0.01" data-field="pm-bottom" value="${margins.bottom ?? 0.75}">
        </div>
        <div class="sn-form-row">
            <label>Left</label>
            <input type="number" step="0.01" data-field="pm-left" value="${margins.left ?? 0.7}">
        </div>
        <div class="sn-form-row">
            <label>Right</label>
            <input type="number" step="0.01" data-field="pm-right" value="${margins.right ?? 0.7}">
        </div>
        <div class="sn-form-row">
            <label>Header</label>
            <input type="number" step="0.01" data-field="pm-header" value="${margins.header ?? 0.3}">
        </div>
        <div class="sn-form-row">
            <label>Footer</label>
            <input type="number" step="0.01" data-field="pm-footer" value="${margins.footer ?? 0.3}">
        </div>
        <div class="sn-form-tip">Unit: inch</div>
    `;

    await SN.Utils.modal({
        titleKey: 'action.print.modal.m001.title', title: 'Margins',
        content: html,
        confirmKey: 'action.print.modal.m001.confirm', confirmText: 'OK',
        cancelKey: 'action.print.modal.m001.cancel', cancelText: 'Cancel',
        onConfirm: (bodyEl) => {
            const read = (field, fallback) => parsePositiveNumber(bodyEl.querySelector(`[data-field="${field}"]`)?.value, fallback);
            const next = {
                top: read('pm-top', margins.top ?? 0.75),
                bottom: read('pm-bottom', margins.bottom ?? 0.75),
                left: read('pm-left', margins.left ?? 0.7),
                right: read('pm-right', margins.right ?? 0.7),
                header: read('pm-header', margins.header ?? 0.3),
                footer: read('pm-footer', margins.footer ?? 0.3)
            };
            SN.Print.setPageMargins(next);
            return true;
        }
    });
}

export function setPageOrientation(orientation) {
    this.SN.Print.setPageSetup({ orientation });
}

export function setPaperSize(paperName) {
    this.SN.Print.setPageSetup({ paperName });
}

export function setPrintScale(scale) {
    this.SN.Print.setPageSetup({ scale, fitToWidth: 0, fitToHeight: 0 });
}

export async function setPrintScaleDialog() {
    const SN = this.SN;
    const settings = SN.Print.ensureSettings(SN.activeSheet);
    const scale = settings.pageSetup?.scale ?? 100;
    const html = `
        <div class="sn-form-row">
            <label>Scale</label>
            <input type="number" step="1" data-field="print-scale" value="${scale}">
        </div>
        <div class="sn-form-tip">Range: 10-400</div>
    `;

    await SN.Utils.modal({
        titleKey: 'action.print.modal.m002.title', title: 'Print Scale',
        content: html,
        confirmKey: 'action.print.modal.m002.confirm', confirmText: 'OK',
        cancelKey: 'action.print.modal.m002.cancel', cancelText: 'Cancel',
        onConfirm: (bodyEl) => {
            const val = Number(bodyEl.querySelector('[data-field="print-scale"]')?.value);
            if (!Number.isFinite(val) || val < 10 || val > 400) {
                SN.Utils.toast(SN.t('action.print.toast.pleaseEnterANumberBetween'));
                return false;
            }
            SN.Print.setPageSetup({ scale: val, fitToWidth: 0, fitToHeight: 0 });
            return true;
        }
    });
}

export function setPrintAreaFromSelection() {
    const sheet = this.SN.activeSheet;
    const area = sheet?.activeAreas?.[sheet.activeAreas.length - 1];
    if (!area) return this.SN.Utils.toast(this.SN.t('action.print.toast.pleaseSelectAPrintArea'));
    this.SN.Print.setPrintArea(area);
    this.SN.Utils.toast(this.SN.t('action.print.toast.printAreaSet'));
}

export function clearPrintArea() {
    this.SN.Print.clearPrintArea();
    this.SN.Utils.toast(this.SN.t('action.print.toast.printAreaCleared'));
}

export async function setPrintTitles() {
    const SN = this.SN;
    const settings = SN.Print.ensureSettings(SN.activeSheet);
    const titles = settings.printTitles || {};
    const html = `
        <div class="sn-form-row">
            <label>Rows to Repeat</label>
            <input type="text" data-field="print-rows" placeholder="e.g. 1:1" value="${escapeAttr(titles.rows ?? '')}">
        </div>
        <div class="sn-form-row">
            <label>Columns to Repeat</label>
            <input type="text" data-field="print-cols" placeholder="e.g. A:A" value="${escapeAttr(titles.cols ?? '')}">
        </div>
    `;

    await SN.Utils.modal({
        titleKey: 'action.print.modal.m003.title', title: 'Print Titles',
        content: html,
        confirmKey: 'action.print.modal.m003.confirm', confirmText: 'OK',
        cancelKey: 'action.print.modal.m003.cancel', cancelText: 'Cancel',
        onConfirm: (bodyEl) => {
            const rows = bodyEl.querySelector('[data-field="print-rows"]')?.value?.trim();
            const cols = bodyEl.querySelector('[data-field="print-cols"]')?.value?.trim();
            SN.Print.setPrintTitles({
                rows: rows || null,
                cols: cols || null
            });
            return true;
        }
    });
}

export async function setPageBackground() {
    const SN = this.SN;
    const settings = SN.Print.ensureSettings(SN.activeSheet);
    const bg = settings.pageBackground?.url ?? '';
    const html = `
        <div class="sn-form-row">
            <label>Background Image URL</label>
            <input type="text" data-field="print-bg" placeholder="https://..." value="${escapeAttr(bg)}">
        </div>
        <div class="sn-form-tip">Background applies to preview and print</div>
    `;

    await SN.Utils.modal({
        titleKey: 'action.print.modal.m004.title', title: 'Background Image',
        content: html,
        confirmKey: 'action.print.modal.m004.confirm', confirmText: 'OK',
        cancelKey: 'action.print.modal.m004.cancel', cancelText: 'Cancel',
        onConfirm: (bodyEl) => {
            const url = bodyEl.querySelector('[data-field="print-bg"]')?.value?.trim();
            SN.Print.setSettings({ pageBackground: url ? { url } : null });
            return true;
        }
    });
}

export async function setHeaderFooter() {
    const SN = this.SN;
    const settings = SN.Print.ensureSettings(SN.activeSheet);
    const headerFooter = settings.headerFooter || {};
    const html = `
        <div class="sn-form-row">
            <label>Header</label>
            <input type="text" data-field="print-header" value="${escapeAttr(headerFooter.header ?? '')}">
        </div>
        <div class="sn-form-row">
            <label>Footer</label>
            <input type="text" data-field="print-footer" value="${escapeAttr(headerFooter.footer ?? '')}">
        </div>
        <div class="sn-form-tip">Header/Footer applies to preview and print</div>
    `;

    await SN.Utils.modal({
        titleKey: 'action.print.modal.m005.title', title: 'Header/Footer',
        content: html,
        confirmKey: 'action.print.modal.m005.confirm', confirmText: 'OK',
        cancelKey: 'action.print.modal.m005.cancel', cancelText: 'Cancel',
        onConfirm: (bodyEl) => {
            const header = bodyEl.querySelector('[data-field="print-header"]')?.value ?? '';
            const footer = bodyEl.querySelector('[data-field="print-footer"]')?.value ?? '';
            SN.Print.setSettings({ headerFooter: { header, footer } });
            return true;
        }
    });
}
