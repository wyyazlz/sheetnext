// 格式操作模块
import { modifyDecimalPlaces } from './helpers.js';
import { formatValue } from '../core/Cell/helpers.js';

/**
 * 检测选区是否为整行/整列选择
 * @param {Object} region - 选区 { s: {r, c}, e: {r, c} }
 * @param {Object} sheet - 工作表
 * @returns {{ isFullRow: boolean, isFullCol: boolean, rows: number[], cols: number[] }}
 */
function detectFullRowCol(region, sheet) {
    const { s, e } = region;
    const isFullRow = s.c === 0 && e.c === sheet.colCount - 1;
    const isFullCol = s.r === 0 && e.r === sheet.rowCount - 1;

    const rows = [];
    const cols = [];

    if (isFullRow) {
        for (let r = s.r; r <= e.r; r++) rows.push(r);
    }
    if (isFullCol) {
        for (let c = s.c; c <= e.c; c++) cols.push(c);
    }

    return { isFullRow, isFullCol, rows, cols };
}

const TEXT_TO_COLUMNS_DELIMITER_OPTIONS = [
    { key: 'tab', labelKey: 'action.format.content.textToColumns.delimiter.tab', value: '\t' },
    { key: 'semicolon', labelKey: 'action.format.content.textToColumns.delimiter.semicolon', value: ';' },
    { key: 'comma', labelKey: 'action.format.content.textToColumns.delimiter.comma', value: ',' },
    { key: 'space', labelKey: 'action.format.content.textToColumns.delimiter.space', value: ' ' }
];

function normalizeTextToColumnsArea(area) {
    if (!area?.s || !area?.e) return null;
    return {
        s: {
            r: Math.min(area.s.r, area.e.r),
            c: Math.min(area.s.c, area.e.c)
        },
        e: {
            r: Math.max(area.s.r, area.e.r),
            c: Math.max(area.s.c, area.e.c)
        }
    };
}

function buildTextToColumnsDialogContent(defaultDestination, SN) {
    const t = (key, params) => SN.t(key, params);
    return `
        <div class="sn-text-to-columns-dialog">
            <div class="sn-form-group">
                <label>${t('action.format.content.textToColumns.delimiter.label')}</label>
                <div class="sn-text-to-columns-delimiters">
                    ${TEXT_TO_COLUMNS_DELIMITER_OPTIONS.map((item, idx) => `
                        <label class="sn-delimiter-item">
                            <input type="checkbox" data-field="delimiter" data-key="${item.key}"${idx === 2 ? ' checked' : ''}>
                            <span>${t(item.labelKey)}</span>
                        </label>
                    `).join('')}
                    <label class="sn-delimiter-item">
                        <input type="checkbox" data-field="delimiter" data-key="other">
                        <span>${t('action.format.content.textToColumns.delimiter.other')}</span>
                    </label>
                    <input class="sn-text-to-columns-other-input" type="text" data-field="other-delimiter" maxlength="1" placeholder="${t('action.format.content.textToColumns.otherPlaceholder')}">
                </div>
            </div>
            <div class="sn-form-group sn-form-group-toggle">
                <label class="sn-inline-toggle">
                    <input type="checkbox" data-field="merge-consecutive">
                    <span>${t('action.format.content.textToColumns.mergeConsecutive')}</span>
                </label>
            </div>
            <div class="sn-form-group">
                <label>${t('action.format.content.textToColumns.textQualifierLabel')}</label>
                <select data-field="text-qualifier">
                    <option value="&quot;">${t('action.format.content.textToColumns.textQualifier.doubleQuote')}</option>
                    <option value="'">${t('action.format.content.textToColumns.textQualifier.singleQuote')}</option>
                    <option value="none">${t('action.format.content.textToColumns.textQualifier.none')}</option>
                </select>
            </div>
            <div class="sn-form-group">
                <label>${t('action.format.content.textToColumns.destinationLabel')}</label>
                <input type="text" data-field="destination" value="${defaultDestination}" placeholder="${t('action.format.content.textToColumns.destinationPlaceholder')}">
            </div>
        </div>
    `;
}

function collectTextToColumnsDelimiters(bodyEl, Utils, SN) {
    const checkedKeys = Array.from(bodyEl.querySelectorAll('[data-field="delimiter"]:checked'))
        .map(input => input.dataset.key);

    const delimiters = TEXT_TO_COLUMNS_DELIMITER_OPTIONS
        .filter(item => checkedKeys.includes(item.key))
        .map(item => item.value);

    if (checkedKeys.includes('other')) {
        const otherDelimiterInput = bodyEl.querySelector('[data-field="other-delimiter"]');
        const rawOther = otherDelimiterInput?.value ?? '';
        if (!rawOther) {
            Utils.toast(SN.t('action.format.toast.msg001'));
            return null;
        }
        delimiters.push(rawOther[0]);
    }

    if (delimiters.length === 0) {
        Utils.toast(SN.t('action.format.toast.msg002'));
        return null;
    }
    return delimiters;
}

export function decimalPlaces(type) {
    const sheet = this.SN.activeSheet;
    sheet.eachCells(sheet.activeAreas, (r, c) => {
        const cell = sheet.getCell(r, c);
        if (cell.type == 'number') {
            cell.numFmt = modifyDecimalPlaces(cell.numFmt, type);
        }
    });
    this.SN._recordChange({ type: 'format', action: 'decimalPlaces', sheet, areas: sheet.activeAreas, formatType: type });
    this.SN._r();
}

export function areasNumFmt(fmtStr) {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;
    // 收集行/列
    const styledRows = new Set();
    const styledCols = new Set();

    areas.forEach(region => {
        if (typeof region === 'string') region = sheet.rangeStrToNum(region);
        const info = detectFullRowCol(region, sheet);

        if (info.isFullRow && !info.isFullCol) {
            info.rows.forEach(r => styledRows.add(r));
        } else if (info.isFullCol && !info.isFullRow) {
            info.cols.forEach(c => styledCols.add(c));
        }
    });

    // 应用行样式
    styledRows.forEach(r => {
        sheet.getRow(r).numFmt = fmtStr;
    });

    // 应用列样式
    styledCols.forEach(c => {
        sheet.getCol(c).numFmt = fmtStr;
    });

    // 对非整行整列的区域，应用到单元格
    sheet.eachCells(areas, (r, c, range) => {
        const info = detectFullRowCol({ s: range.s, e: range.e }, sheet);
        if ((info.isFullRow && !info.isFullCol) || (info.isFullCol && !info.isFullRow)) return;
        const cell = sheet.getCell(r, c);
        if (fmtStr == '@') cell.type = 'string';
        if (fmtStr == '') cell.type = undefined;
        cell.numFmt = fmtStr;
    });

    this.SN._recordChange({ type: 'format', action: 'areasNumFmt', sheet, areas, fmtStr });
    this.SN._r();
}

export function clearAreaFormat(type = 'all') {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;
    // 收集行/列
    const styledRows = new Set();
    const styledCols = new Set();

    areas.forEach(region => {
        if (typeof region === 'string') region = sheet.rangeStrToNum(region);
        const info = detectFullRowCol(region, sheet);

        if (info.isFullRow && !info.isFullCol) {
            info.rows.forEach(r => styledRows.add(r));
        } else if (info.isFullCol && !info.isFullRow) {
            info.cols.forEach(c => styledCols.add(c));
        }
    });

    // 清除行样式
    styledRows.forEach(r => {
        const row = sheet.getRow(r);
        if (type === 'style' || type === 'all') {
            row.style = null;
        }
        if (type === 'border') {
            row.border = null;
        }
    });

    // 清除列样式
    styledCols.forEach(c => {
        const col = sheet.getCol(c);
        if (type === 'style' || type === 'all') {
            col.style = null;
        }
        if (type === 'border') {
            col.border = null;
        }
    });

    // 对非整行整列的区域，应用到单元格
    sheet.eachCells(areas, (r, c, range) => {
        const info = detectFullRowCol({ s: range.s, e: range.e }, sheet);
        if ((info.isFullRow && !info.isFullCol) || (info.isFullCol && !info.isFullRow)) return;
        const cell = sheet.getCell(r, c);
        if (type == 'style') {
            cell.style = {};
        } else if (type == 'value') {
            cell.editVal = "";
        } else if (type == 'border') {
            cell.border = null;
        } else {
            cell.style = {};
            cell.editVal = "";
            cell.border = null;
        }
    });

    this.SN._recordChange({ type: 'format', action: 'clearAreaFormat', sheet, areas, clearType: type });
    this.SN._r();
}

// 文本转数值
export function textToNumber() {
    const sheet = this.SN.activeSheet;
    sheet.eachCells(sheet.activeAreas, (r, c) => {
        const cell = sheet.getCell(r, c);
        const val = cell.calcVal;
        if (typeof val === 'string' && val !== '' && !isNaN(Number(val))) {
            cell.editVal = Number(val);
        }
    });
    this.SN._recordChange({ type: 'format', action: 'textToNumber', sheet, areas: sheet.activeAreas });
    this.SN._r();
}

// 数值转文本
export function numberToText() {
    const sheet = this.SN.activeSheet;
    sheet.eachCells(sheet.activeAreas, (r, c) => {
        const cell = sheet.getCell(r, c);
        const val = cell.calcVal;
        if (typeof val === 'number') {
            cell.numFmt = '@';
        }
    });
    this.SN._recordChange({ type: 'format', action: 'numberToText', sheet, areas: sheet.activeAreas });
    this.SN._r();
}

// 大写转小写
export function toLowercase() {
    const sheet = this.SN.activeSheet;
    sheet.eachCells(sheet.activeAreas, (r, c) => {
        const cell = sheet.getCell(r, c);
        const val = cell._editVal;
        if (typeof val === 'string' && !val.startsWith('=')) {
            cell.editVal = val.toLowerCase();
        }
    });
    this.SN._recordChange({ type: 'format', action: 'toLowercase', sheet, areas: sheet.activeAreas });
    this.SN._r();
}

// 小写转大写
export function toUppercase() {
    const sheet = this.SN.activeSheet;
    sheet.eachCells(sheet.activeAreas, (r, c) => {
        const cell = sheet.getCell(r, c);
        const val = cell._editVal;
        if (typeof val === 'string' && !val.startsWith('=')) {
            cell.editVal = val.toUpperCase();
        }
    });
    this.SN._recordChange({ type: 'format', action: 'toUppercase', sheet, areas: sheet.activeAreas });
    this.SN._r();
}

// 去除空格
export function trimSpaces() {
    const sheet = this.SN.activeSheet;
    sheet.eachCells(sheet.activeAreas, (r, c) => {
        const cell = sheet.getCell(r, c);
        const val = cell._editVal;
        if (typeof val === 'string' && !val.startsWith('=')) {
            cell.editVal = val.trim().replace(/\s+/g, ' ');
        }
    });
    this.SN._recordChange({ type: 'format', action: 'trimSpaces', sheet, areas: sheet.activeAreas });
    this.SN._r();
}

// Text to Columns
export function textToColumns() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const area = normalizeTextToColumnsArea(sheet.activeAreas?.[0]);

    if (!area) {
        Utils.toast(SN.t('action.format.toast.msg003'));
        return;
    }

    if (area.s.c !== area.e.c) {
        Utils.toast(SN.t('action.format.toast.msg004'));
        return;
    }

    if (sheet.areaHaveMerge(area)) {
        Utils.toast(SN.t('action.format.toast.msg005'));
        return;
    }

    const defaultDestination = Utils.cellNumToStr(area.s);
    Utils.modal({
        titleKey: 'action.format.modal.m001.title',
        title: SN.t('action.format.modal.m001.title'),
        content: buildTextToColumnsDialogContent(defaultDestination, SN),
        maxWidth: '520px',
        confirmKey: 'action.format.modal.m001.confirm',
        confirmText: SN.t('action.format.modal.m001.confirm'),
        cancelKey: 'action.format.modal.m001.cancel',
        cancelText: SN.t('action.format.modal.m001.cancel'),
        onOpen: (bodyEl) => {
            const otherDelimiterInput = bodyEl.querySelector('[data-field="other-delimiter"]');
            const otherDelimiterCheckbox = bodyEl.querySelector('[data-key="other"]');
            otherDelimiterInput?.addEventListener('input', () => {
                if (otherDelimiterInput.value) otherDelimiterCheckbox.checked = true;
            });
            bodyEl.querySelector('[data-field="destination"]')?.focus();
        },
        onConfirm: (bodyEl) => {
            const delimiters = collectTextToColumnsDelimiters(bodyEl, Utils, SN);
            if (!delimiters) return false;

            const destinationInput = bodyEl.querySelector('[data-field="destination"]');
            const destinationValue = destinationInput?.value?.replace(/\$/g, '').trim();
            const destination = Utils.cellStrToNum(destinationValue);
            if (!destination || destination.r < 0 || destination.c < 0) {
                Utils.toast(SN.t('action.format.toast.a1'));
                return false;
            }

            const textQualifier = bodyEl.querySelector('[data-field="text-qualifier"]')?.value ?? '"';
            const mergeConsecutive = bodyEl.querySelector('[data-field="merge-consecutive"]')?.checked === true;

            let result;
            try {
                result = sheet.textToColumns(area, {
                    delimiters,
                    destination,
                    textQualifier,
                    treatConsecutiveDelimitersAsOne: mergeConsecutive
                });
            } catch (e) {
                Utils.toast(e?.message || SN.t('action.format.toast.textToColumnsFailed'));
                return false;
            }

            if (!result?.success) {
                if (!result?.canceled) Utils.toast(SN.t('action.format.toast.msg007'));
                return false;
            }

            this.SN._r();
            Utils.toast(SN.t('action.format.toast.textToColumnsDone', { rowCount: result.rowCount, columnCount: result.columnCount }));
            return true;
        }
    });
}

export function flashFill() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const Utils = SN.Utils;
    const area = normalizeTextToColumnsArea(sheet.activeAreas?.[0]);

    if (!area) {
        Utils.toast(SN.t('action.format.toast.msg008'));
        return;
    }

    if (area.s.c !== area.e.c) {
        Utils.toast(SN.t('action.format.toast.msg009'));
        return;
    }

    if (sheet.areaHaveMerge(area)) {
        Utils.toast(SN.t('action.format.toast.msg010'));
        return;
    }

    let result;
    try {
        result = sheet.flashFill(area);
    } catch (e) {
        Utils.toast(e?.message || SN.t('action.format.toast.flashFillFailed'));
        return;
    }

    if (!result?.success) {
        if (!result?.canceled) Utils.toast(SN.t('action.format.toast.msg011'));
        return;
    }

    this.SN._r();
    if (result.filledCount > 0) {
        Utils.toast(SN.t('action.format.toast.flashFillDone', { filledCount: result.filledCount }));
    } else {
        Utils.toast(SN.t('action.format.toast.msg012'));
    }
}

// Custom number format dialog
export function customNumFmtDialog() {
    const t = (key, params) => this.SN.t(key, params);
    const sheet = this.SN.activeSheet;
    const { r, c } = sheet.activeCell;
    const cell = sheet.getCell(r, c);
    const currentFmt = cell.numFmt || '';
    const testValue = typeof cell.calcVal === 'number' ? cell.calcVal : 1234.567;

    // 格式化预览函数
    const formatPreview = (fmt, val) => {
        if (!fmt) return String(val);
        try {
            const result = formatValue(val, fmt);
            return result?.text ?? String(val);
        } catch {
            return String(val);
        }
    };

    const quickFormats = [
        { fmt: '#,##0', descKey: 'action.format.content.customNumFmt.quick.integer' },
        { fmt: '0.00', descKey: 'action.format.content.customNumFmt.quick.decimal2' },
        { fmt: '0.00%', descKey: 'action.format.content.customNumFmt.quick.percent' },
        { fmt: '¥#,##0.00', descKey: 'action.format.content.customNumFmt.quick.currency' },
        { fmt: 'yyyy/m/d', descKey: 'action.format.content.customNumFmt.quick.date' },
        { fmt: 'h:mm:ss', descKey: 'action.format.content.customNumFmt.quick.time' }
    ];

    const content = `
        <div class="sn-numfmt-dialog">
            <div class="sn-numfmt-row">
                <label>${t('action.format.content.customNumFmt.currentLabel')}</label>
                <code class="sn-numfmt-current">${currentFmt || t('action.format.content.customNumFmt.none')}</code>
            </div>
            <div class="sn-numfmt-row">
                <label>${t('action.format.content.customNumFmt.inputLabel')}</label>
                <input type="text" class="sn-numfmt-input" value="${currentFmt}" placeholder="${t('action.format.content.customNumFmt.inputPlaceholder')}">
            </div>
            <div class="sn-numfmt-row">
                <label>${t('action.format.content.customNumFmt.previewLabel')}</label>
                <span class="sn-numfmt-preview">${formatPreview(currentFmt, testValue)}</span>
            </div>
            <details class="sn-numfmt-help" open>
                <summary>${t('action.format.content.customNumFmt.placeholdersTitle')}</summary>
                <div class="sn-numfmt-placeholders">
                    <div><code>0</code> ${t('action.format.content.customNumFmt.placeholder.zero')}</div>
                    <div><code>#</code> ${t('action.format.content.customNumFmt.placeholder.hash')}</div>
                    <div><code>?</code> ${t('action.format.content.customNumFmt.placeholder.question')}</div>
                    <div><code>.</code> ${t('action.format.content.customNumFmt.placeholder.dot')}</div>
                    <div><code>,</code> ${t('action.format.content.customNumFmt.placeholder.comma')}</div>
                    <div><code>%</code> ${t('action.format.content.customNumFmt.placeholder.percent')}</div>
                    <div><code>@</code> ${t('action.format.content.customNumFmt.placeholder.at')}</div>
                    <div><code>yyyy</code> ${t('action.format.content.customNumFmt.placeholder.year')}</div>
                    <div><code>m/mm</code> ${t('action.format.content.customNumFmt.placeholder.month')}</div>
                    <div><code>d/dd</code> ${t('action.format.content.customNumFmt.placeholder.day')}</div>
                    <div><code>h:mm:ss</code> ${t('action.format.content.customNumFmt.placeholder.time')}</div>
                    <div><code>[Red]</code> ${t('action.format.content.customNumFmt.placeholder.red')}</div>
                </div>
            </details>
            <div class="sn-numfmt-quick">
                <label>${t('action.format.content.customNumFmt.quickLabel')}</label>
                <div class="sn-numfmt-tags">
                    ${quickFormats.map((item) => `<span class="sn-numfmt-tag" data-fmt="${item.fmt}" title="${t(item.descKey)}">${item.fmt}</span>`).join('')}
                </div>
            </div>
        </div>
    `;

    this.SN.Utils.modal({
        titleKey: 'action.format.modal.m002.title',
        title: t('action.format.modal.m002.title'),
        maxWidth: '420px',
        content,
        confirmKey: 'action.format.modal.m002.confirm',
        confirmText: t('action.format.modal.m002.confirm'),
        cancelKey: 'action.format.modal.m002.cancel',
        cancelText: t('action.format.modal.m002.cancel'),
        onOpen: (bodyEl) => {
            const input = bodyEl.querySelector('.sn-numfmt-input');
            const preview = bodyEl.querySelector('.sn-numfmt-preview');
            const tags = bodyEl.querySelectorAll('.sn-numfmt-tag');

            // 实时预览
            input.addEventListener('input', () => {
                preview.textContent = formatPreview(input.value, testValue);
            });

            // 快捷标签点击
            tags.forEach(tag => {
                tag.addEventListener('click', () => {
                    input.value = tag.dataset.fmt;
                    preview.textContent = formatPreview(input.value, testValue);
                    input.focus();
                });
            });

            input.focus();
            input.select();
        },
        onConfirm: (bodyEl) => {
            const fmt = bodyEl.querySelector('.sn-numfmt-input').value;
            this.areasNumFmt(fmt);
        }
    });
}
