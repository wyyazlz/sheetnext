function escapeAttr(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

export function printPreview() {
    if (this.SN.Canvas.inputEditing) this.SN.Canvas.updInputValue();
    return this.SN.Print.preview();
}

export async function exportPdf() {
    await this.SN.Print.print();
}

export function pageBreakPreview() {
    this.SN.Print.toggleShowPageBreaks(true);
    this.SN.Utils.toast(this.SN.t('core.print.print.pageBreakHint'));
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
        const cell = this.SN.activeSheet.activeCell;
        const key = cell.r === 0 && cell.c === 0 ? 'cannotInsertPageBreakAtOrigin' : 'aPageBreakAlreadyExists';
        this.SN.Utils.toast(this.SN.t(`action.print.toast.${key}`));
        return;
    }
    this.SN.Print.toggleShowPageBreaks(true);
    const setup = this.SN.Print.ensureSettings().pageSetup;
    const key = setup.fitToWidth > 0 || setup.fitToHeight > 0 ? 'pageBreakInsertedWhileFitting' : 'pageBreakInserted';
    this.SN.Utils.toast(this.SN.t(`action.print.toast.${key}`));
}

export function togglePrintGridlines(checked) {
    this.SN.Print.toggleGridlines(!!checked);
}

export function togglePrintHeadings(checked) {
    this.SN.Print.toggleHeadings(!!checked);
}

export function setPageMarginsPreset(preset) {
    this.SN.Print.setPageMarginsPreset(preset);
    this.SN.Utils.toast(this.SN.t('action.print.toast.marginsUpdated'));
}

export async function setPageMarginsDialog() {
    const SN = this.SN;
    const settings = SN.Print.ensureSettings(SN.activeSheet);
    const margins = settings.pageMargins || {};
    const fields = ['top', 'bottom', 'left', 'right', 'header', 'footer'];
    const html = fields.map(field => `<label class="sn-form-row">
        <span>${escapeAttr(SN.t(`core.print.print.margins.${field}`))}</span>
        <input type="number" min="0" step="0.01" required data-field="${field}" value="${(margins[field] * 2.54).toFixed(2)}">
    </label>`).join('') + `<div class="sn-form-tip">${escapeAttr(SN.t('core.print.print.marginUnit'))}</div>`;

    await SN.Utils.modal({
        titleKey: 'action.print.modal.m001.title', title: 'Margins',
        content: html,
        confirmKey: 'action.print.modal.m001.confirm', confirmText: 'OK',
        cancelKey: 'action.print.modal.m001.cancel', cancelText: 'Cancel',
        onConfirm: (bodyEl) => {
            const inputs = Array.from(bodyEl.querySelectorAll('input'));
            if (inputs.some(input => !input.reportValidity())) return false;
            const next = Object.fromEntries(inputs.map(input => [input.dataset.field, Number(input.value) / 2.54]));
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
            <label>${escapeAttr(SN.t('core.print.print.scalePercent'))}</label>
            <input type="number" step="1" data-field="print-scale" value="${scale}">
        </div>
        <div class="sn-form-tip">${escapeAttr(SN.t('core.print.print.scaleHint'))}</div>
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
            <label>${escapeAttr(SN.t('core.print.print.titleRows'))}</label>
            <input type="text" data-field="print-rows" placeholder="1:1" value="${escapeAttr(titles.rows ?? '')}">
        </div>
        <div class="sn-form-row">
            <label>${escapeAttr(SN.t('core.print.print.titleCols'))}</label>
            <input type="text" data-field="print-cols" placeholder="A:A" value="${escapeAttr(titles.cols ?? '')}">
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
            <label>${escapeAttr(SN.t('core.print.print.backgroundUrl'))}</label>
            <input type="text" data-field="print-bg" placeholder="https://..." value="${escapeAttr(bg)}">
        </div>
        <div class="sn-form-tip">${escapeAttr(SN.t('core.print.print.backgroundHint'))}</div>
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
            <label>${escapeAttr(SN.t('core.print.print.header'))}</label>
            <input type="text" data-field="print-header" value="${escapeAttr(headerFooter.header ?? '')}">
        </div>
        <div class="sn-form-row">
            <label>${escapeAttr(SN.t('core.print.print.footer'))}</label>
            <input type="text" data-field="print-footer" value="${escapeAttr(headerFooter.footer ?? '')}">
        </div>
        <div class="sn-form-tip">${escapeAttr(SN.t('core.print.print.headerFooterHint'))}</div>
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
