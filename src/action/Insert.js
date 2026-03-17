const OPEN_DRAWING_MESSAGE = 'Drawing and chart insertion is disabled in SheetNext Open Edition.';

function _notify(SN, message = OPEN_DRAWING_MESSAGE) {
    if (SN?.Utils?.toast) SN.Utils.toast(message);
    return false;
}

function _normalizeSparklineKey(type) {
    const raw = String(type || '').trim().toLowerCase();
    if (!raw) return 'line';
    if (raw === 'stacked' || raw === 'winloss' || raw === 'win-loss' || raw === 'win_loss') return 'win_loss';
    if (raw === 'column' || raw === 'line') return raw;
    return 'line';
}

function _normalizeSparklineType(type) {
    const key = _normalizeSparklineKey(type);
    return key === 'win_loss' ? 'stacked' : key;
}

function _normalizeSheetName(name) {
    return name.includes(' ') ? `'${name}'` : name;
}

function _buildRangeRef(sheet, sr, sc, er, ec) {
    const sheetName = _normalizeSheetName(sheet.name);
    const start = sheet.Utils.cellNumToStr({ r: sr, c: sc });
    const end = sheet.Utils.cellNumToStr({ r: er, c: ec });
    return `${sheetName}!${start}:${end}`;
}

function _resolveSparklineRange(SN, defaultSheet, rangeInput) {
    const input = String(rangeInput || '').trim();
    if (!input) throw SN.t('action.insert.content.sparkline.rangeRequired');

    let sourceSheet = defaultSheet;
    let rangeRef = input;

    if (input.includes('!')) {
        const [sheetName, rangePart] = input.split('!');
        const cleanSheetName = sheetName.replace(/^'|'$/g, '');
        const targetSheet = SN.getSheet(cleanSheetName);
        if (!targetSheet) throw SN.t('action.insert.content.sparkline.sheetNotFound', { name: cleanSheetName });
        sourceSheet = targetSheet;
        rangeRef = rangePart;
    }

    let rangeObj = null;
    try {
        rangeObj = sourceSheet.rangeStrToNum(rangeRef);
    } catch (_error) {
        throw SN.t('action.insert.content.sparkline.rangeInvalid');
    }
    if (!rangeObj) throw SN.t('action.insert.content.sparkline.rangeInvalid');

    const formula = _buildRangeRef(sourceSheet, rangeObj.s.r, rangeObj.s.c, rangeObj.e.r, rangeObj.e.c);
    return { formula };
}

function _resolveSparklineLocation(SN, defaultSheet, locationInput) {
    const input = String(locationInput || '').trim();
    if (!input) throw SN.t('action.insert.content.sparkline.locationRequired');

    let targetSheet = defaultSheet;
    let cellRef = input;

    if (input.includes('!')) {
        const [sheetName, cellPart] = input.split('!');
        const cleanSheetName = sheetName.replace(/^'|'$/g, '');
        const target = SN.getSheet(cleanSheetName);
        if (!target) throw SN.t('action.insert.content.sparkline.sheetNotFound', { name: cleanSheetName });
        targetSheet = target;
        cellRef = cellPart;
    }

    if (cellRef.includes(':')) throw SN.t('action.insert.content.sparkline.locationSingleCell');

    const cellObj = targetSheet.Utils.cellStrToNum(cellRef);
    if (!cellObj) throw SN.t('action.insert.content.sparkline.locationInvalid');

    return { targetSheet, cellObj };
}

export async function insertComment() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    if (!sheet?.Comment) return false;

    const activeCell = sheet.activeCell;
    const cellRef = SN.Utils.cellNumToStr(activeCell);
    const existingComment = sheet.Comment.get(cellRef);
    const modalTitleKey = existingComment
        ? 'action.insert.modal.m001.titleEdit'
        : 'action.insert.modal.m001.titleInsert';

    const result = await SN.Utils.modal({
        titleKey: modalTitleKey,
        title: SN.t(modalTitleKey),
        maxWidth: '400px',
        content: `
            <div style="margin-bottom: 10px;">
                <label style="display: block; margin-bottom: 5px;">Cell: ${cellRef}</label>
            </div>
            <div style="margin-bottom: 10px;">
                <label style="display: block; margin-bottom: 5px;">Note:</label>
                <textarea data-field="comment-text" style="width: 100%; border: 1px solid #ccc; height: 80px; resize: none; padding: 8px; box-sizing: border-box;">${existingComment?.text || ''}</textarea>
            </div>
        `,
        confirmKey: 'action.insert.modal.m001.confirm',
        confirmText: 'OK',
        cancelKey: 'action.insert.modal.m001.cancel',
        cancelText: 'Cancel'
    });
    if (!result) return false;

    const textarea = result.bodyEl.querySelector('[data-field="comment-text"]');
    const text = textarea?.value?.trim();

    if (!text) {
        if (existingComment) {
            sheet.Comment.remove(cellRef);
            SN._r();
        }
        return true;
    }

    if (existingComment) {
        existingComment.text = text;
    } else {
        sheet.Comment.add({ cellRef, text });
    }

    SN._r();
    return true;
}

export async function insertAreaScreenshot() {
    return _notify(this?.SN);
}

export async function downloadActiveDrawingImage() {
    return _notify(this?.SN);
}

export async function insertImages() {
    return _notify(this?.SN);
}

export function insertTextBox() {
    return _notify(this?.SN);
}

export async function openChartModal() {
    return _notify(this?.SN);
}

export function insertChart() {
    return _notify(this?.SN);
}

export function insertSparkline(type) {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const areas = sheet.activeAreas;

    if (!areas || areas.length === 0) {
        SN.Utils.toast(SN.t('action.insert.toast.msg002'));
        return false;
    }

    if (!sheet.Sparkline) {
        SN.Utils.toast(SN.t('action.insert.toast.msg003'));
        return false;
    }

    const area = areas[areas.length - 1];
    const formula = _buildRangeRef(sheet, area.s.r, area.s.c, area.e.r, area.e.c);
    const targetCell = sheet.activeCell;

    sheet.Sparkline.add({
        cellRef: { r: targetCell.r, c: targetCell.c },
        formula,
        type: _normalizeSparklineType(type)
    });

    SN._r();
    return true;
}

export async function openSparklineDialog(type) {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const areas = sheet.activeAreas;
    const area = areas && areas.length ? areas[areas.length - 1] : null;
    const defaultRange = area ? `${sheet.name}!${SN.Utils.rangeNumToStr(area)}` : '';
    const defaultLocation = sheet.Utils.cellNumToStr(sheet.activeCell);
    const defaultTypeKey = _normalizeSparklineKey(type);
    const typeLabels = {
        line: 'action.insert.content.sparkline.type.line',
        column: 'action.insert.content.sparkline.type.column',
        win_loss: 'action.insert.content.sparkline.type.winLoss'
    };

    const typeOptions = Object.entries(typeLabels).map(([key, labelKey]) => (
        `<option value="${key}" ${key === defaultTypeKey ? 'selected' : ''}>${SN.t(labelKey)}</option>`
    )).join('');

    try {
        await SN.Utils.modal({
            titleKey: 'action.insert.modal.m003.title',
            title: SN.t('action.insert.modal.m003.title'),
            maxWidth: '420px',
            content: `
                <div style="display:flex;flex-direction:column;gap:12px;">
                    <div>
                        <div style="margin-bottom:6px;font-size:13px;color:#333;">${SN.t('action.insert.content.sparkline.rangeLabel')}</div>
                        <input type="text" data-field="sparkline-range" value="${defaultRange}" style="width:100%;" />
                    </div>
                    <div>
                        <div style="margin-bottom:6px;font-size:13px;color:#333;">${SN.t('action.insert.content.sparkline.locationLabel')}</div>
                        <input type="text" data-field="sparkline-location" value="${defaultLocation}" style="width:100%;" />
                    </div>
                    <div>
                        <div style="margin-bottom:6px;font-size:13px;color:#333;">${SN.t('action.insert.content.sparkline.typeLabel')}</div>
                        <select data-field="sparkline-type" style="width:100%;">${typeOptions}</select>
                    </div>
                </div>
            `,
            confirmKey: 'action.insert.modal.m003.confirm',
            confirmText: SN.t('action.insert.modal.m003.confirm'),
            cancelKey: 'action.insert.modal.m003.cancel',
            cancelText: SN.t('action.insert.modal.m003.cancel'),
            onConfirm: (bodyEl) => {
                const rangeInput = bodyEl.querySelector('[data-field="sparkline-range"]').value.trim();
                const locationInput = bodyEl.querySelector('[data-field="sparkline-location"]').value.trim();
                const typeValue = bodyEl.querySelector('[data-field="sparkline-type"]').value;

                const { formula } = _resolveSparklineRange(SN, sheet, rangeInput);
                const { targetSheet, cellObj } = _resolveSparklineLocation(SN, sheet, locationInput);

                if (!targetSheet.Sparkline) throw SN.t('action.insert.content.sparkline.moduleNotReady');

                targetSheet.Sparkline.add({
                    cellRef: { r: cellObj.r, c: cellObj.c },
                    formula,
                    type: _normalizeSparklineType(typeValue)
                });

                SN.activeSheet = targetSheet;
                SN._r();
            }
        });
        return true;
    } catch (_error) {
        return false;
    }
}

export async function insertHyperlink() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const activeCell = sheet.activeCell;
    const cell = sheet.getCell(activeCell.r, activeCell.c);
    const existingLink = cell.hyperlink;
    const currentText = cell.editVal || '';
    const sheetNames = SN.sheets.map(s => s.name);
    const sheetOptions = sheetNames.map(n => `<option value="${n}">${n}</option>`).join('');

    let existingSheet = sheetNames[0] || '';
    let existingCellRef = 'A1';
    if (existingLink?.location) {
        const match = existingLink.location.match(/^(.+?)!(.+)$/);
        if (match) {
            existingSheet = match[1].replace(/^'|'$/g, '');
            existingCellRef = match[2];
        }
    }

    const isExternal = !existingLink || existingLink.target;
    const t = (key, params = {}) => SN.t(key, params);
    const modalTitleKey = existingLink
        ? 'action.insert.modal.m004.titleEdit'
        : 'action.insert.modal.m004.titleInsert';
    const modalCancelKey = existingLink
        ? 'action.insert.modal.m004.cancelRemove'
        : 'action.insert.modal.m004.cancel';

    try {
        await SN.Utils.modal({
            titleKey: modalTitleKey,
            title: SN.t(modalTitleKey),
            maxWidth: '400px',
            content: `
                <style>
                    .sn-hyperlink-form { display: flex; flex-direction: column; gap: 12px; min-width: 320px; }
                    .sn-hyperlink-form label { display: block; font-size: 13px; color: #333; margin-bottom: 4px; }
                    .sn-hyperlink-form input, .sn-hyperlink-form select {
                        width: 100%; height: 32px; padding: 0 8px; border: 1px solid #d0d0d0;
                        border-radius: 4px; font-size: 13px; box-sizing: border-box;
                    }
                    .sn-hyperlink-form input:focus, .sn-hyperlink-form select:focus {
                        outline: none; border-color: #4f7cf6; box-shadow: 0 0 0 2px rgba(79,124,246,0.15);
                    }
                    .sn-hyperlink-row { display: flex; gap: 8px; }
                    .sn-hyperlink-row > div { flex: 1; }
                </style>
                <div class="sn-hyperlink-form">
                    <div>
                        <label>${t('action.insert.modal.m004.label.displayText')}</label>
                        <input type="text" data-field="link-text" placeholder="${t('action.insert.modal.m004.placeholder.displayText')}" value="${currentText}">
                    </div>
                    <div>
                        <label>${t('action.insert.modal.m004.label.linkType')}</label>
                        <select data-field="link-type">
                            <option value="external" ${isExternal ? 'selected' : ''}>${t('action.insert.modal.m004.option.external')}</option>
                            <option value="internal" ${!isExternal ? 'selected' : ''}>${t('action.insert.modal.m004.option.internal')}</option>
                        </select>
                    </div>
                    <div data-field="link-external" style="display:${isExternal ? 'block' : 'none'}">
                        <label>${t('action.insert.modal.m004.label.address')}</label>
                        <input type="text" data-field="link-target" placeholder="${t('action.insert.modal.m004.placeholder.address')}" value="${existingLink?.target || ''}">
                    </div>
                    <div data-field="link-internal" style="display:${!isExternal ? 'block' : 'none'}">
                        <label>${t('action.insert.modal.m004.label.targetLocation')}</label>
                        <div class="sn-hyperlink-row">
                            <div>
                                <select data-field="link-sheet">${sheetOptions}</select>
                            </div>
                            <div>
                                <input type="text" data-field="link-cell" placeholder="${t('action.insert.modal.m004.placeholder.cell')}" value="${existingCellRef}">
                            </div>
                        </div>
                    </div>
                    <div>
                        <label>${t('action.insert.modal.m004.label.screenTip')}</label>
                        <input type="text" data-field="link-tooltip" placeholder="${t('action.insert.modal.m004.placeholder.screenTip')}" value="${existingLink?.tooltip || ''}">
                    </div>
                </div>
            `,
            confirmKey: 'action.insert.modal.m004.confirm',
            confirmText: 'OK',
            cancelKey: modalCancelKey,
            cancelText: SN.t(modalCancelKey),
            onOpen(bodyEl) {
                bodyEl.querySelector('[data-field="link-sheet"]').value = existingSheet;
                const typeSelect = bodyEl.querySelector('[data-field="link-type"]');
                const externalDiv = bodyEl.querySelector('[data-field="link-external"]');
                const internalDiv = bodyEl.querySelector('[data-field="link-internal"]');
                typeSelect.onchange = () => {
                    externalDiv.style.display = typeSelect.value === 'external' ? 'block' : 'none';
                    internalDiv.style.display = typeSelect.value === 'internal' ? 'block' : 'none';
                };
            },
            onConfirm(bodyEl) {
                const linkType = bodyEl.querySelector('[data-field="link-type"]').value;
                const displayText = bodyEl.querySelector('[data-field="link-text"]').value.trim();
                const tooltip = bodyEl.querySelector('[data-field="link-tooltip"]').value.trim();
                const target = bodyEl.querySelector('[data-field="link-target"]').value.trim();
                const sheetName = bodyEl.querySelector('[data-field="link-sheet"]').value;
                const cellAddr = bodyEl.querySelector('[data-field="link-cell"]').value.trim().toUpperCase() || 'A1';

                if (linkType === 'external' && !target) {
                    SN.Utils.toast(SN.t('action.insert.toast.pleaseEnterALinkAddress'));
                    return false;
                }

                const hyperlink = {};
                if (linkType === 'external') {
                    hyperlink.target = target;
                } else {
                    hyperlink.location = sheetName.includes(' ') ? `'${sheetName}'!${cellAddr}` : `${sheetName}!${cellAddr}`;
                }
                if (tooltip) hyperlink.tooltip = tooltip;

                cell.hyperlink = hyperlink;
                cell.editVal = displayText || target || hyperlink.location;
                SN._r();
            }
        });
        return true;
    } catch (_error) {
        return false;
    }
}
