function _normalizeArea(area) {
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

function _resolveArea(sheet) {
    const area = _normalizeArea(sheet.activeAreas?.[0]);
    if (area) return area;
    return { s: { ...sheet.activeCell }, e: { ...sheet.activeCell } };
}

function _isNumericColumn(sheet, area, colIndex, hasHeaders = true) {
    const start = hasHeaders ? area.s.r + 1 : area.s.r;
    for (let r = start; r <= area.e.r; r++) {
        const value = sheet.getCell(r, colIndex).calcVal;
        if (value === '' || value === null || value === undefined) continue;
        if (typeof value === 'number' && Number.isFinite(value)) return true;
    }
    return false;
}

function _buildColumns(sheet, area, SN, hasHeaders = true) {
    const columns = [];
    const headerRow = area.s.r;
    for (let c = area.s.c; c <= area.e.c; c++) {
        const colRef = sheet.Utils.numToChar(c);
        const headerText = hasHeaders
            ? String(sheet.getCell(headerRow, c).showVal ?? '').trim()
            : '';
        const title = headerText
            ? `${headerText} (${colRef})`
            : `${SN.t('dataTools.columnFallback')} ${colRef}`;
        columns.push({
            col: c,
            ref: colRef,
            title,
            isNumeric: _isNumericColumn(sheet, area, c, hasHeaders)
        });
    }
    return columns;
}

function _buildFunctionOptions(SN) {
    return [
        { key: 'sum', text: SN.t('dataTools.function.sum') },
        { key: 'count', text: SN.t('dataTools.function.count') },
        { key: 'countNums', text: SN.t('dataTools.function.countNums') },
        { key: 'average', text: SN.t('dataTools.function.average') },
        { key: 'max', text: SN.t('dataTools.function.max') },
        { key: 'min', text: SN.t('dataTools.function.min') },
        { key: 'product', text: SN.t('dataTools.function.product') },
        { key: 'stdev', text: SN.t('dataTools.function.stdev') },
        { key: 'stdevp', text: SN.t('dataTools.function.stdevp') },
        { key: 'var', text: SN.t('dataTools.function.var') },
        { key: 'varp', text: SN.t('dataTools.function.varp') }
    ];
}

const DV_TYPE_OPTIONS = [
    { key: 'any', labelKey: 'dataValidation.type.any' },
    { key: 'whole', labelKey: 'dataValidation.type.whole' },
    { key: 'decimal', labelKey: 'dataValidation.type.decimal' },
    { key: 'list', labelKey: 'dataValidation.type.list' },
    { key: 'date', labelKey: 'dataValidation.type.date' },
    { key: 'time', labelKey: 'dataValidation.type.time' },
    { key: 'textLength', labelKey: 'dataValidation.type.textLength' },
    { key: 'custom', labelKey: 'dataValidation.type.custom' }
];

const DV_OPERATOR_OPTIONS = [
    { key: 'between', labelKey: 'dataValidation.operator.between' },
    { key: 'notBetween', labelKey: 'dataValidation.operator.notBetween' },
    { key: 'equal', labelKey: 'dataValidation.operator.equal' },
    { key: 'notEqual', labelKey: 'dataValidation.operator.notEqual' },
    { key: 'greaterThan', labelKey: 'dataValidation.operator.greaterThan' },
    { key: 'lessThan', labelKey: 'dataValidation.operator.lessThan' },
    { key: 'greaterThanOrEqual', labelKey: 'dataValidation.operator.greaterThanOrEqual' },
    { key: 'lessThanOrEqual', labelKey: 'dataValidation.operator.lessThanOrEqual' }
];

function _buildDialogContent(columns, SN, options = {}) {
    const funcOptions = _buildFunctionOptions(SN);
    const fnKey = options.function ?? 'sum';
    const groupBy = options.groupBy ?? columns[0]?.col ?? 0;
    const summaryBelow = options.summaryBelowData !== false;

    const groupOptionsHtml = columns
        .map(col => `<option value="${col.col}"${col.col === groupBy ? ' selected' : ''}>${col.title}</option>`)
        .join('');

    const targetHtml = columns
        .map(col => {
            const checked = options.defaultTargets?.includes(col.col) ? ' checked' : '';
            return `
                <label>
                    <input type="checkbox" data-field="subtotal-target" value="${col.col}"${checked}>
                    <span>${col.title}</span>
                </label>
            `;
        })
        .join('');

    return `
        <div class="sn-subtotal-dialog">
            <div class="sn-form-group">
                <label>${SN.t('dataTools.subtotal.groupField')}</label>
                <select data-field="subtotal-group">${groupOptionsHtml}</select>
            </div>
            <div class="sn-form-group">
                <label>${SN.t('dataTools.subtotal.function')}</label>
                <select data-field="subtotal-function">
                    ${funcOptions.map(item => `<option value="${item.key}"${item.key === fnKey ? ' selected' : ''}>${item.text}</option>`).join('')}
                </select>
            </div>
            <div class="sn-form-group">
                <label>${SN.t('dataTools.subtotal.addTo')}</label>
                <div class="sn-subtotal-targets">${targetHtml}</div>
            </div>
            <div class="sn-form-group">
                <label>
                    <input type="checkbox" data-field="subtotal-has-headers" checked>
                    <span>${SN.t('dataTools.subtotal.hasHeaders')}</span>
                </label>
                <label>
                    <input type="checkbox" data-field="subtotal-replace" checked>
                    <span>${SN.t('dataTools.subtotal.replaceExisting')}</span>
                </label>
                <label>
                    <input type="checkbox" data-field="subtotal-summary-below"${summaryBelow ? ' checked' : ''}>
                    <span>${SN.t('dataTools.subtotal.summaryBelow')}</span>
                </label>
                <label>
                    <input type="checkbox" data-field="subtotal-grand-total" checked>
                    <span>${SN.t('dataTools.subtotal.grandTotal')}</span>
                </label>
            </div>
        </div>
    `;
}

function _refreshTargetDisabledState(bodyEl) {
    const groupSelect = bodyEl.querySelector('[data-field="subtotal-group"]');
    const groupCol = Number(groupSelect?.value);
    const targetInputs = bodyEl.querySelectorAll('[data-field="subtotal-target"]');
    targetInputs.forEach(input => {
        const col = Number(input.value);
        const disabled = col === groupCol;
        input.disabled = disabled;
        if (disabled) input.checked = false;
    });
}

function _collectTargets(bodyEl) {
    return Array.from(bodyEl.querySelectorAll('[data-field="subtotal-target"]:checked'))
        .map(input => Number(input.value))
        .filter(col => Number.isInteger(col) && col >= 0);
}

function _normalizeRangesInput(sheet, text, SN) {
    const entries = String(text ?? '')
        .split(/[\n,;]+/)
        .map(item => item.trim())
        .filter(Boolean);
    if (entries.length === 0) return [];

    const ranges = [];
    entries.forEach(item => {
        const range = sheet.rangeStrToNum(item);
        if (!range?.s || !range?.e) {
            throw new Error(SN.t('dataTools.consolidate.invalidRange', { range: item }));
        }
        ranges.push(item);
    });
    return ranges;
}

function _buildConsolidateDialogContent(defaultRangesText, defaultDestination, SN) {
    const fnOptions = _buildFunctionOptions(SN);
    return `
        <div class="sn-consolidate-dialog">
            <div class="sn-form-group">
                <label>${SN.t('dataTools.consolidate.function')}</label>
                <select data-field="consolidate-function">
                    ${fnOptions.map(item => `<option value="${item.key}"${item.key === 'sum' ? ' selected' : ''}>${item.text}</option>`).join('')}
                </select>
            </div>
            <div class="sn-form-group">
                <label>${SN.t('dataTools.consolidate.sourceRanges')}</label>
                <textarea data-field="consolidate-ranges" rows="6" placeholder="A1:C10&#10;E1:G10">${defaultRangesText}</textarea>
            </div>
            <div class="sn-form-group">
                <label>${SN.t('dataTools.consolidate.destination')}</label>
                <input type="text" data-field="consolidate-destination" value="${defaultDestination}" placeholder="${SN.t('dataTools.consolidate.destinationPlaceholder')}">
            </div>
            <div class="sn-form-group sn-form-group-toggle">
                <label class="sn-inline-toggle">
                    <input type="checkbox" data-field="consolidate-skip-blanks" checked>
                    <span>${SN.t('dataTools.consolidate.skipBlanks')}</span>
                </label>
            </div>
        </div>
    `;
}

function _getActiveValidationRule(sheet) {
    const active = sheet.activeCell;
    return sheet.getCell(active.r, active.c).dataValidation || null;
}

function _buildDataValidationDialogContent(rule, SN) {
    const currentType = rule?.type || 'any';
    const currentOperator = rule?.operator || 'between';
    const listSource = Array.isArray(rule?.formula1) ? rule.formula1.join(',') : '';
    const formula1 = !Array.isArray(rule?.formula1) && rule?.formula1 != null ? String(rule.formula1) : '';
    const formula2 = rule?.formula2 != null ? String(rule.formula2) : '';

    return `
        <div class="sn-data-validation-dialog">
            <div class="sn-data-validation-grid">
                <div class="sn-form-group">
                    <label>${SN.t('dataValidation.allow')}</label>
                    <select data-field="dv-type">
                        ${DV_TYPE_OPTIONS.map(item => `<option value="${item.key}"${item.key === currentType ? ' selected' : ''}>${SN.t(item.labelKey)}</option>`).join('')}
                    </select>
                </div>
                <div class="sn-form-group" data-role="dv-operator-group">
                    <label>${SN.t('dataValidation.data')}</label>
                    <select data-field="dv-operator">
                        ${DV_OPERATOR_OPTIONS.map(item => `<option value="${item.key}"${item.key === currentOperator ? ' selected' : ''}>${SN.t(item.labelKey)}</option>`).join('')}
                    </select>
                </div>
            </div>
            <div class="sn-form-group" data-role="dv-list-source-group">
                <label>${SN.t('dataValidation.listSource')}</label>
                <input type="text" data-field="dv-list-source" value="${listSource}" placeholder="${SN.t('dataValidation.listSourcePlaceholder')}">
            </div>
            <div class="sn-data-validation-grid">
                <div class="sn-form-group" data-role="dv-formula1-group">
                    <label data-role="dv-formula1-label">${SN.t('dataValidation.formula1.default')}</label>
                    <input type="text" data-field="dv-formula1" value="${formula1}">
                </div>
                <div class="sn-form-group" data-role="dv-formula2-group">
                    <label data-role="dv-formula2-label">${SN.t('dataValidation.formula2.default')}</label>
                    <input type="text" data-field="dv-formula2" value="${formula2}">
                </div>
            </div>
            <label class="sn-data-validation-toggle">
                <input type="checkbox" data-field="dv-allow-blank"${rule?.allowBlank !== false ? ' checked' : ''}>
                <span>${SN.t('dataValidation.allowBlank')}</span>
            </label>
            <label class="sn-data-validation-toggle" data-role="dv-dropdown-group">
                <input type="checkbox" data-field="dv-show-dropdown"${rule?.showDropDown !== false ? ' checked' : ''}>
                <span>${SN.t('dataValidation.inCellDropdown')}</span>
            </label>
            <label class="sn-data-validation-toggle">
                <input type="checkbox" data-field="dv-show-input"${rule?.showInputMessage !== false ? ' checked' : ''}>
                <span>${SN.t('dataValidation.showInputPrompt')}</span>
            </label>
            <div class="sn-form-group sn-data-validation-fields" data-role="dv-input-settings">
                <input type="text" data-field="dv-prompt-title" value="${rule?.promptTitle ?? ''}" placeholder="${SN.t('dataValidation.promptTitlePlaceholder')}">
                <textarea data-field="dv-prompt" rows="2" placeholder="${SN.t('dataValidation.promptContentPlaceholder')}">${rule?.prompt ?? ''}</textarea>
            </div>
            <label class="sn-data-validation-toggle">
                <input type="checkbox" data-field="dv-show-error"${rule?.showErrorMessage !== false ? ' checked' : ''}>
                <span>${SN.t('dataValidation.showErrorAlert')}</span>
            </label>
            <div class="sn-form-group sn-data-validation-fields" data-role="dv-error-settings">
                <select data-field="dv-error-style">
                    <option value="stop"${(rule?.errorStyle || 'stop') === 'stop' ? ' selected' : ''}>${SN.t('dataValidation.errorStyle.stop')}</option>
                    <option value="warning"${rule?.errorStyle === 'warning' ? ' selected' : ''}>${SN.t('dataValidation.errorStyle.warning')}</option>
                    <option value="information"${rule?.errorStyle === 'information' ? ' selected' : ''}>${SN.t('dataValidation.errorStyle.information')}</option>
                </select>
                <input type="text" data-field="dv-error-title" value="${rule?.errorTitle ?? ''}" placeholder="${SN.t('dataValidation.errorTitlePlaceholder')}">
                <textarea data-field="dv-error" rows="2" placeholder="${SN.t('dataValidation.errorContentPlaceholder')}">${rule?.error ?? ''}</textarea>
            </div>
        </div>
    `;
}

function _updateDataValidationDialogState(bodyEl, SN) {
    const type = bodyEl.querySelector('[data-field="dv-type"]')?.value || 'any';
    const operator = bodyEl.querySelector('[data-field="dv-operator"]')?.value || 'between';
    const operatorGroup = bodyEl.querySelector('[data-role="dv-operator-group"]');
    const listSourceGroup = bodyEl.querySelector('[data-role="dv-list-source-group"]');
    const formula1Group = bodyEl.querySelector('[data-role="dv-formula1-group"]');
    const formula2Group = bodyEl.querySelector('[data-role="dv-formula2-group"]');
    const dropdownGroup = bodyEl.querySelector('[data-role="dv-dropdown-group"]');
    const inputSettings = bodyEl.querySelector('[data-role="dv-input-settings"]');
    const errorSettings = bodyEl.querySelector('[data-role="dv-error-settings"]');
    const showInputCheckbox = bodyEl.querySelector('[data-field="dv-show-input"]');
    const showErrorCheckbox = bodyEl.querySelector('[data-field="dv-show-error"]');
    const formula1Label = bodyEl.querySelector('[data-role="dv-formula1-label"]');
    const formula2Label = bodyEl.querySelector('[data-role="dv-formula2-label"]');

    const needsOperator = ['whole', 'decimal', 'textLength', 'date', 'time'].includes(type);
    const isList = type === 'list';
    const isCustom = type === 'custom';
    const needsFormula = needsOperator || isCustom;
    const needsFormula2 = needsOperator && ['between', 'notBetween'].includes(operator);

    if (operatorGroup) operatorGroup.style.display = needsOperator ? '' : 'none';
    if (listSourceGroup) listSourceGroup.style.display = isList ? '' : 'none';
    if (formula1Group) formula1Group.style.display = needsFormula ? '' : 'none';
    if (formula2Group) formula2Group.style.display = needsFormula2 ? '' : 'none';
    if (dropdownGroup) dropdownGroup.style.display = isList ? '' : 'none';
    if (inputSettings) inputSettings.style.display = showInputCheckbox?.checked === false ? 'none' : '';
    if (errorSettings) errorSettings.style.display = showErrorCheckbox?.checked === false ? 'none' : '';

    if (!formula1Label || !formula2Label) return;
    if (type === 'date') {
        formula1Label.textContent = SN.t('dataValidation.formula1.date');
        formula2Label.textContent = SN.t('dataValidation.formula2.date');
    } else if (type === 'time') {
        formula1Label.textContent = SN.t('dataValidation.formula1.time');
        formula2Label.textContent = SN.t('dataValidation.formula2.time');
    } else if (type === 'textLength') {
        formula1Label.textContent = SN.t('dataValidation.formula1.textLength');
        formula2Label.textContent = SN.t('dataValidation.formula2.textLength');
    } else if (isCustom) {
        formula1Label.textContent = SN.t('dataValidation.formula1.custom');
        formula2Label.textContent = SN.t('dataValidation.formula2.custom');
    } else {
        formula1Label.textContent = SN.t('dataValidation.formula1.default');
        formula2Label.textContent = SN.t('dataValidation.formula2.default');
    }
}

function _buildDataValidationRuleFromDialog(bodyEl, SN) {
    const type = bodyEl.querySelector('[data-field="dv-type"]')?.value || 'any';
    if (type === 'any') return null;

    const rule = {
        type,
        allowBlank: bodyEl.querySelector('[data-field="dv-allow-blank"]')?.checked !== false,
        showInputMessage: bodyEl.querySelector('[data-field="dv-show-input"]')?.checked !== false,
        showErrorMessage: bodyEl.querySelector('[data-field="dv-show-error"]')?.checked !== false,
        promptTitle: bodyEl.querySelector('[data-field="dv-prompt-title"]')?.value?.trim() || '',
        prompt: bodyEl.querySelector('[data-field="dv-prompt"]')?.value?.trim() || '',
        errorStyle: bodyEl.querySelector('[data-field="dv-error-style"]')?.value || 'stop',
        errorTitle: bodyEl.querySelector('[data-field="dv-error-title"]')?.value?.trim() || '',
        error: bodyEl.querySelector('[data-field="dv-error"]')?.value?.trim() || ''
    };

    if (type === 'list') {
        const listSource = bodyEl.querySelector('[data-field="dv-list-source"]')?.value || '';
        const values = listSource
            .split(/[,\n，;；]/)
            .map(item => item.trim())
            .filter(Boolean);
        if (!values.length) throw new Error(SN.t('dataValidation.errors.listSourceRequired'));
        rule.formula1 = values;
        rule.showDropDown = bodyEl.querySelector('[data-field="dv-show-dropdown"]')?.checked !== false;
        return rule;
    }

    if (type === 'custom') {
        const formula = bodyEl.querySelector('[data-field="dv-formula1"]')?.value?.trim() || '';
        if (!formula) throw new Error(SN.t('dataValidation.errors.customFormulaRequired'));
        rule.formula1 = formula;
        return rule;
    }

    if (['whole', 'decimal', 'textLength', 'date', 'time'].includes(type)) {
        const operator = bodyEl.querySelector('[data-field="dv-operator"]')?.value || 'between';
        const formula1 = bodyEl.querySelector('[data-field="dv-formula1"]')?.value?.trim() || '';
        const formula2 = bodyEl.querySelector('[data-field="dv-formula2"]')?.value?.trim() || '';
        if (!formula1) throw new Error(SN.t('dataValidation.errors.conditionRequired'));
        rule.operator = operator;
        rule.formula1 = formula1;
        if (['between', 'notBetween'].includes(operator)) {
            if (!formula2) throw new Error(SN.t('dataValidation.errors.rangeConditionRequired'));
            rule.formula2 = formula2;
        }
    }

    return rule;
}

export function subtotal() {
    const SN = this.SN;
    const sheet = this.SN.activeSheet;
    const Utils = this.SN.Utils;
    const area = _resolveArea(sheet);

    if (sheet.areaHaveMerge(area)) {
        Utils.toast(SN.t('dataTools.subtotal.errors.mergedCells'));
        return;
    }

    const columns = _buildColumns(sheet, area, SN, true);
    if (columns.length < 2) {
        Utils.toast(SN.t('dataTools.subtotal.errors.minColumns'));
        return;
    }

    const defaultTargets = columns
        .filter(col => col.isNumeric && col.col !== columns[0].col)
        .map(col => col.col);
    if (defaultTargets.length === 0) {
        columns.forEach(col => {
            if (col.col !== columns[0].col) defaultTargets.push(col.col);
        });
    }

    Utils.modal({
        titleKey: 'dataTools.subtotal.title',
        title: 'Subtotal',
        content: _buildDialogContent(columns, SN, {
            groupBy: columns[0].col,
            function: 'sum',
            defaultTargets,
            summaryBelowData: sheet.outlinePr?.summaryBelow !== false
        }),
        maxWidth: '520px',
        confirmKey: 'common.ok',
        confirmText: 'OK',
        cancelKey: 'common.cancel',
        cancelText: 'Cancel',
        onOpen: (bodyEl) => {
            const groupSelect = bodyEl.querySelector('[data-field="subtotal-group"]');
            _refreshTargetDisabledState(bodyEl);
            groupSelect?.addEventListener('change', () => _refreshTargetDisabledState(bodyEl));
        },
        onConfirm: (bodyEl) => {
            const groupBy = Number(bodyEl.querySelector('[data-field="subtotal-group"]')?.value);
            const functionName = bodyEl.querySelector('[data-field="subtotal-function"]')?.value || 'sum';
            const subtotalColumns = _collectTargets(bodyEl);
            const hasHeaders = bodyEl.querySelector('[data-field="subtotal-has-headers"]')?.checked !== false;
            const replaceExisting = bodyEl.querySelector('[data-field="subtotal-replace"]')?.checked !== false;
            const summaryBelowData = bodyEl.querySelector('[data-field="subtotal-summary-below"]')?.checked !== false;
            const withGrandTotal = bodyEl.querySelector('[data-field="subtotal-grand-total"]')?.checked !== false;

            if (!Number.isInteger(groupBy)) {
                Utils.toast(SN.t('dataTools.subtotal.errors.groupFieldRequired'));
                return false;
            }
            if (!subtotalColumns.length) {
                Utils.toast(SN.t('dataTools.subtotal.errors.subtotalColumnRequired'));
                return false;
            }

            let result;
            try {
                result = sheet.subtotal(area, {
                    groupBy,
                    function: functionName,
                    subtotalColumns,
                    hasHeaders,
                    replaceExisting,
                    summaryBelowData,
                    withGrandTotal
                });
            } catch (e) {
                Utils.toast(e?.message || SN.t('dataTools.subtotal.errors.executeFailed'));
                return false;
            }

            if (!result?.success) return false;
            this.SN._r();
            const hasGrandTotal = result.grandTotalRow != null
                ? SN.t('dataTools.subtotal.completed.grandTotalSuffix')
                : '';
            Utils.toast(SN.t('dataTools.subtotal.completed.message', {
                    groupCount: result.groupCount,
                    insertedRows: result.insertedSubtotalRows,
                    grandTotal: hasGrandTotal
                }));
            return true;
        }
    });
}

export function consolidate() {
    const SN = this.SN;
    const sheet = this.SN.activeSheet;
    const Utils = this.SN.Utils;
    const rangesText = (sheet.activeAreas || [])
        .map(area => sheet.Utils.rangeNumToStr(_normalizeArea(area)))
        .filter(Boolean)
        .join('\n');
    const defaultRangesText = rangesText || sheet.Utils.rangeNumToStr({ s: sheet.activeCell, e: sheet.activeCell });
    const defaultDestination = sheet.Utils.cellNumToStr(sheet.activeCell);

    Utils.modal({
        titleKey: 'dataTools.consolidate.title',
        title: 'Consolidate',
        maxWidth: '520px',
        content: _buildConsolidateDialogContent(defaultRangesText, defaultDestination, SN),
        confirmKey: 'common.ok',
        confirmText: 'OK',
        cancelKey: 'common.cancel',
        cancelText: 'Cancel',
        onConfirm: (bodyEl) => {
            const fnName = bodyEl.querySelector('[data-field="consolidate-function"]')?.value || 'sum';
            const rangesRaw = bodyEl.querySelector('[data-field="consolidate-ranges"]')?.value;
            const destinationRaw = bodyEl.querySelector('[data-field="consolidate-destination"]')?.value?.trim();
            const skipBlanks = bodyEl.querySelector('[data-field="consolidate-skip-blanks"]')?.checked !== false;

            let sourceRanges;
            try {
                sourceRanges = _normalizeRangesInput(sheet, rangesRaw, SN);
            } catch (e) {
                Utils.toast(e?.message || SN.t('dataTools.consolidate.errors.invalidSourceRange'));
                return false;
            }
            if (!sourceRanges.length) {
                Utils.toast(SN.t('dataTools.consolidate.errors.sourceRequired'));
                return false;
            }

            let result;
            try {
                result = sheet.consolidate(sourceRanges, {
                    function: fnName,
                    destination: destinationRaw,
                    skipBlanks
                });
            } catch (e) {
                Utils.toast(e?.message || SN.t('dataTools.consolidate.errors.executeFailed'));
                return false;
            }

            if (!result?.success) return false;
            this.SN._r();
            Utils.toast(SN.t('dataTools.consolidate.completed', { count: result.writtenCount }));
            return true;
        }
    });
}

export function dataValidationDialog() {
    const sheet = this.SN.activeSheet;
    const SN = this.SN;
    const Utils = this.SN.Utils;
    const currentRule = _getActiveValidationRule(sheet);

    Utils.modal({
        titleKey: 'dataValidation.title',
        title: SN.t('dataValidation.title'),
        maxWidth: '520px',
        content: _buildDataValidationDialogContent(currentRule, SN),
        confirmKey: 'common.ok',
        confirmText: SN.t('common.ok'),
        cancelKey: 'common.cancel',
        cancelText: SN.t('common.cancel'),
        onOpen: (bodyEl) => {
            const typeSelect = bodyEl.querySelector('[data-field="dv-type"]');
            const operatorSelect = bodyEl.querySelector('[data-field="dv-operator"]');
            const showInputCheckbox = bodyEl.querySelector('[data-field="dv-show-input"]');
            const showErrorCheckbox = bodyEl.querySelector('[data-field="dv-show-error"]');
            _updateDataValidationDialogState(bodyEl, SN);
            typeSelect?.addEventListener('change', () => _updateDataValidationDialogState(bodyEl, SN));
            operatorSelect?.addEventListener('change', () => _updateDataValidationDialogState(bodyEl, SN));
            showInputCheckbox?.addEventListener('change', () => _updateDataValidationDialogState(bodyEl, SN));
            showErrorCheckbox?.addEventListener('change', () => _updateDataValidationDialogState(bodyEl, SN));
        },
        onConfirm: (bodyEl) => {
            let rule = null;
            try {
                rule = _buildDataValidationRuleFromDialog(bodyEl, SN);
            } catch (e) {
                Utils.toast(e?.message || SN.t('dataValidation.errors.invalidConfig'));
                return false;
            }

            const areas = sheet.activeAreas?.length
                ? sheet.activeAreas
                : [{ s: sheet.activeCell, e: sheet.activeCell }];
            let result;
            try {
                result = rule
                    ? sheet.setDataValidation(areas, rule)
                    : sheet.clearDataValidation(areas);
            } catch (e) {
                Utils.toast(e?.message || SN.t('dataValidation.errors.setFailed'));
                return false;
            }
            if (!result?.success) return false;
            this.SN._r();
            const actionText = rule
                ? SN.t('dataValidation.action.set')
                : SN.t('dataValidation.action.clear');
            Utils.toast(SN.t('dataValidation.toast.completed', { action: actionText, count: result.cellCount }));
            return true;
        }
    });
}

export function clearDataValidation() {
    const sheet = this.SN.activeSheet;
    const Utils = this.SN.Utils;
    const areas = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];

    let result;
    try {
        result = sheet.clearDataValidation(areas);
    } catch (e) {
        Utils.toast(e?.message || this.SN.t('dataValidation.errors.clearFailed'));
        return;
    }
    if (!result?.success) return;

    this.SN._r();
    Utils.toast(this.SN.t('dataValidation.toast.cleared', {
        count: result.cellCount
    }));
}
