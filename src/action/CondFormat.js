/**
 * 条件格式操作模块
 * 处理条件格式菜单的所有交互逻辑
 */

// 格式预设
const FORMAT_PRESETS = [
    { labelKey: 'action.condFormat.content.preset.lightRedText', fill: '#FFC7CE', font: '#9C0006' },
    { labelKey: 'action.condFormat.content.preset.yellowText', fill: '#FFEB9C', font: '#9C6500' },
    { labelKey: 'action.condFormat.content.preset.greenText', fill: '#C6EFCE', font: '#006100' },
    { labelKey: 'action.condFormat.content.preset.lightRed', fill: '#FFC7CE', font: null },
    { labelKey: 'action.condFormat.content.preset.yellow', fill: '#FFEB9C', font: null },
    { labelKey: 'action.condFormat.content.preset.green', fill: '#C6EFCE', font: null }
];

// 操作符标签
const OPERATOR_LABELS = {
    greaterThan: 'action.condFormat.content.operator.greaterThan',
    lessThan: 'action.condFormat.content.operator.lessThan',
    between: 'action.condFormat.content.operator.between',
    equal: 'action.condFormat.content.operator.equal',
    notEqual: 'action.condFormat.content.operator.notEqual',
    greaterThanOrEqual: 'action.condFormat.content.operator.greaterThanOrEqual',
    lessThanOrEqual: 'action.condFormat.content.operator.lessThanOrEqual'
};

// 时间周期标签
const TIME_PERIOD_LABELS = {
    today: 'action.condFormat.content.timePeriod.today',
    yesterday: 'action.condFormat.content.timePeriod.yesterday',
    tomorrow: 'action.condFormat.content.timePeriod.tomorrow',
    last7Days: 'action.condFormat.content.timePeriod.last7Days',
    lastWeek: 'action.condFormat.content.timePeriod.lastWeek',
    thisWeek: 'action.condFormat.content.timePeriod.thisWeek',
    nextWeek: 'action.condFormat.content.timePeriod.nextWeek',
    lastMonth: 'action.condFormat.content.timePeriod.lastMonth',
    thisMonth: 'action.condFormat.content.timePeriod.thisMonth',
    nextMonth: 'action.condFormat.content.timePeriod.nextMonth'
};

function getFormatPresets(SN) {
    return FORMAT_PRESETS.map((preset) => ({
        ...preset,
        label: SN.t(preset.labelKey)
    }));
}

function getOperatorLabel(SN, operator) {
    const key = OPERATOR_LABELS[operator];
    return key ? SN.t(key) : operator;
}

function getTimePeriodEntries(SN) {
    return Object.entries(TIME_PERIOD_LABELS).map(([key, labelKey]) => [key, SN.t(labelKey)]);
}

/**
 * 条件格式统一入口
 * @param {string} type - 规则类型
 * @param {string} subType - 子类型或预设
 * @param {string} extra - 额外参数（如颜色）
 */
export async function condFormat(type, subType, extra) {
    const SN = this.SN;
    const sheet = SN.activeSheet;

    // 获取选区字符串
    const areas = sheet.activeAreas;
    if (!areas || areas.length === 0) {
        SN.Utils.toast(SN.t('action.condFormat.toast.msg001'));
        return;
    }
    const sqref = SN.Utils.rangeNumToStr(areas[areas.length - 1]);

    try {
        switch (type) {
            case 'cellIs':
                await showCellIsDialog.call(this, sqref, subType);
                break;

            case 'containsText':
                await showTextDialog.call(this, sqref, 'containsText');
                break;

            case 'timePeriod':
                await showTimePeriodDialog.call(this, sqref);
                break;

            case 'duplicateValues':
                await showDuplicateDialog.call(this, sqref);
                break;

            case 'top10':
                await showTop10Dialog.call(this, sqref, subType);
                break;

            case 'aboveAverage':
                applyAboveAverage.call(this, sqref, subType);
                break;

            case 'dataBar':
                if (subType === 'custom') {
                    await showDataBarDialog.call(this, sqref);
                } else {
                    applyDataBar.call(this, sqref, subType === 'gradient', extra);
                }
                break;

            case 'colorScale':
                if (subType === 'custom') {
                    await showColorScaleDialog.call(this, sqref);
                } else {
                    applyColorScale.call(this, sqref, subType);
                }
                break;

            case 'iconSet':
                if (subType === 'custom') {
                    await showIconSetDialog.call(this, sqref);
                } else {
                    applyIconSet.call(this, sqref, subType);
                }
                break;

            case 'newRule':
                await showNewRuleDialog.call(this, sqref);
                break;

            case 'clearSelection':
                clearSelectionRules.call(this, sqref);
                break;

            case 'clearSheet':
                clearSheetRules.call(this);
                break;

            case 'manage':
                await showManageDialog.call(this);
                break;
        }
    } catch (e) {
        // 用户取消
        if (e !== 'cancel') console.error(SN.t('action.condFormat.content.errorPrefix'), e);
    }
}

// ==================== 弹窗实现 ====================

/**
 * 单元格值比较弹窗
 */
async function showCellIsDialog(sqref, operator) {
    const SN = this.SN;
    const isBetween = operator === 'between';
    const title = getOperatorLabel(SN, operator);
    const presets = getFormatPresets(SN);

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.cellIs.description', { title })}</span>
            </div>
            <div class="sn-cf-form-row">
                <input type="text" class="sn-cf-form-input" data-field="cf-value1" placeholder="${SN.t('action.condFormat.content.cellIs.value1Placeholder')}">
                ${isBetween
                    ? `<span>${SN.t('action.condFormat.content.cellIs.and')}</span><input type="text" class="sn-cf-form-input" data-field="cf-value2" placeholder="${SN.t('action.condFormat.content.cellIs.value2Placeholder')}">`
                    : ''}
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.setTo')}</span>
                <select class="sn-cf-form-select" data-field="cf-preset">
                    ${presets.map((preset, i) => `<option value="${i}">${preset.label}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview">AaBbCc</div>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m001.title',
        title: SN.t('action.condFormat.content.modalTitle', { title }),
        content: html,
        confirmKey: 'action.condFormat.modal.m001.confirm',
        confirmText: SN.t('action.condFormat.modal.m001.confirm'),
        cancelKey: 'action.condFormat.modal.m001.cancel',
        cancelText: SN.t('action.condFormat.modal.m001.cancel'),
        onOpen: (bodyEl) => {
            // 预览实时更新
            const presetSelect = bodyEl.querySelector('[data-field="cf-preset"]');
            const previewEl = bodyEl.querySelector('[data-field="cf-preview"]');
            const updatePreview = () => {
                const preset = presets[presetSelect.value];
                previewEl.style.background = preset.fill;
                previewEl.style.color = preset.font || '#000';
            };
            presetSelect.addEventListener('change', updatePreview);
            updatePreview();
        }
    });
    if (!result) throw 'cancel';

    const body = result.bodyEl;
    const v1 = body.querySelector('[data-field="cf-value1"]')?.value?.trim();
    const v2 = body.querySelector('[data-field="cf-value2"]')?.value?.trim();
    const presetIdx = body.querySelector('[data-field="cf-preset"]')?.value || 0;
    const preset = presets[presetIdx];

    if (!v1) {
        SN.Utils.toast(SN.t('action.condFormat.toast.msg002'));
        throw 'cancel';
    }

    // 应用规则
    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: 'cellIs',
        operator,
        formula1: parseValue(v1),
        formula2: isBetween ? parseValue(v2) : null,
        dxf: {
            fill: { fgColor: preset.fill },
            font: preset.font ? { color: preset.font } : null
        }
    });
    SN._r();
}

/**
 * 文本包含弹窗
 */
async function showTextDialog(sqref, type) {
    const SN = this.SN;
    const presets = getFormatPresets(SN);

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.containsText.description')}</span>
            </div>
            <div class="sn-cf-form-row">
                <input type="text" class="sn-cf-form-input" data-field="cf-text" placeholder="${SN.t('action.condFormat.content.containsText.placeholder')}">
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.setTo')}</span>
                <select class="sn-cf-form-select" data-field="cf-preset">
                    ${presets.map((preset, i) => `<option value="${i}">${preset.label}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview">AaBbCc</div>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m002.title',
        title: SN.t('action.condFormat.content.modalTitle', { title: SN.t('action.condFormat.content.title.containsText') }),
        content: html,
        confirmKey: 'action.condFormat.modal.m002.confirm',
        confirmText: SN.t('action.condFormat.modal.m002.confirm'),
        cancelKey: 'action.condFormat.modal.m002.cancel',
        cancelText: SN.t('action.condFormat.modal.m002.cancel'),
        onOpen: (bodyEl) => {
            const presetSelect = bodyEl.querySelector('[data-field="cf-preset"]');
            const previewEl = bodyEl.querySelector('[data-field="cf-preview"]');
            const updatePreview = () => {
                const preset = presets[presetSelect.value];
                previewEl.style.background = preset.fill;
                previewEl.style.color = preset.font || '#000';
            };
            presetSelect.addEventListener('change', updatePreview);
            updatePreview();
        }
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const text = body.querySelector('[data-field="cf-text"]')?.value?.trim();
    const preset = presets[body.querySelector('[data-field="cf-preset"]').value];

    if (!text) {
        SN.Utils.toast(SN.t('action.condFormat.toast.msg003'));
        throw 'cancel';
    }

    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: 'containsText',
        text,
        dxf: {
            fill: { fgColor: preset.fill },
            font: preset.font ? { color: preset.font } : null
        }
    });
    SN._r();
}

/**
 * 时间周期弹窗
 */
async function showTimePeriodDialog(sqref) {
    const SN = this.SN;
    const periods = getTimePeriodEntries(SN);
    const presets = getFormatPresets(SN);

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.timePeriod.description')}</span>
            </div>
            <div class="sn-cf-form-row">
                <select class="sn-cf-form-select" data-field="cf-period" style="width:100%">
                    ${periods.map(([key, label]) => `<option value="${key}">${label}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.setTo')}</span>
                <select class="sn-cf-form-select" data-field="cf-preset">
                    ${presets.map((preset, i) => `<option value="${i}">${preset.label}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview">AaBbCc</div>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m003.title',
        title: SN.t('action.condFormat.content.modalTitle', { title: SN.t('action.condFormat.content.title.timePeriod') }),
        content: html,
        confirmKey: 'action.condFormat.modal.m003.confirm',
        confirmText: SN.t('action.condFormat.modal.m003.confirm'),
        cancelKey: 'action.condFormat.modal.m003.cancel',
        cancelText: SN.t('action.condFormat.modal.m003.cancel'),
        onOpen: (bodyEl) => {
            const presetSelect = bodyEl.querySelector('[data-field="cf-preset"]');
            const previewEl = bodyEl.querySelector('[data-field="cf-preview"]');
            const updatePreview = () => {
                const preset = presets[presetSelect.value];
                previewEl.style.background = preset.fill;
                previewEl.style.color = preset.font || '#000';
            };
            presetSelect.addEventListener('change', updatePreview);
            updatePreview();
        }
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const period = body.querySelector('[data-field="cf-period"]').value;
    const preset = presets[body.querySelector('[data-field="cf-preset"]').value];

    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: 'timePeriod',
        timePeriod: period,
        dxf: {
            fill: { fgColor: preset.fill },
            font: preset.font ? { color: preset.font } : null
        }
    });
    SN._r();
}

/**
 * 重复值弹窗
 */
async function showDuplicateDialog(sqref) {
    const SN = this.SN;
    const presets = getFormatPresets(SN);

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.duplicate.description')}</span>
            </div>
            <div class="sn-cf-form-row">
                <select class="sn-cf-form-select" data-field="cf-duptype" style="width:100%">
                    <option value="duplicateValues">${SN.t('action.condFormat.content.duplicate.duplicateValues')}</option>
                    <option value="uniqueValues">${SN.t('action.condFormat.content.duplicate.uniqueValues')}</option>
                </select>
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.setTo')}</span>
                <select class="sn-cf-form-select" data-field="cf-preset">
                    ${presets.map((preset, i) => `<option value="${i}">${preset.label}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview">AaBbCc</div>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m004.title',
        title: SN.t('action.condFormat.content.modalTitle', { title: SN.t('action.condFormat.content.title.duplicate') }),
        content: html,
        confirmKey: 'action.condFormat.modal.m004.confirm',
        confirmText: SN.t('action.condFormat.modal.m004.confirm'),
        cancelKey: 'action.condFormat.modal.m004.cancel',
        cancelText: SN.t('action.condFormat.modal.m004.cancel'),
        onOpen: (bodyEl) => {
            const presetSelect = bodyEl.querySelector('[data-field="cf-preset"]');
            const previewEl = bodyEl.querySelector('[data-field="cf-preview"]');
            const updatePreview = () => {
                const preset = presets[presetSelect.value];
                previewEl.style.background = preset.fill;
                previewEl.style.color = preset.font || '#000';
            };
            presetSelect.addEventListener('change', updatePreview);
            updatePreview();
        }
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const dupType = body.querySelector('[data-field="cf-duptype"]').value;
    const preset = presets[body.querySelector('[data-field="cf-preset"]').value];

    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: dupType,
        dxf: {
            fill: { fgColor: preset.fill },
            font: preset.font ? { color: preset.font } : null
        }
    });
    SN._r();
}

/**
 * 前10项弹窗
 */
async function showTop10Dialog(sqref, subType) {
    const SN = this.SN;
    const isBottom = subType.includes('bottom');
    const isPercent = subType.includes('Percent');
    const presets = getFormatPresets(SN);

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.top10.descriptionPrefix', { target: SN.t(isBottom ? 'action.condFormat.content.top10.min' : 'action.condFormat.content.top10.max') })}</span>
                <input type="number" class="sn-cf-form-input" data-field="cf-rank" value="10" min="1" max="1000" style="width:60px">
                <span>${isPercent ? '%' : SN.t('action.condFormat.content.top10.items')}</span>
                <span>${SN.t('action.condFormat.content.top10.setFormat')}</span>
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.setTo')}</span>
                <select class="sn-cf-form-select" data-field="cf-preset">
                    ${presets.map((preset, i) => `<option value="${i}">${preset.label}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview">AaBbCc</div>
            </div>
        </div>
    `;

    const modalTitleKey = isBottom
        ? (isPercent
            ? 'action.condFormat.modal.m005.titleBottomPercent'
            : 'action.condFormat.modal.m005.titleBottomItems')
        : (isPercent
            ? 'action.condFormat.modal.m005.titleTopPercent'
            : 'action.condFormat.modal.m005.titleTopItems');

    const result = await SN.Utils.modal({
        titleKey: modalTitleKey,
        title: SN.t(modalTitleKey),
        content: html,
        confirmKey: 'action.condFormat.modal.m005.confirm',
        cancelKey: 'action.condFormat.modal.m005.cancel',
        onOpen: (bodyEl) => {
            const presetSelect = bodyEl.querySelector('[data-field="cf-preset"]');
            const previewEl = bodyEl.querySelector('[data-field="cf-preview"]');
            const updatePreview = () => {
                const preset = presets[presetSelect.value];
                previewEl.style.background = preset.fill;
                previewEl.style.color = preset.font || '#000';
            };
            presetSelect.addEventListener('change', updatePreview);
            updatePreview();
        }
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const rank = parseInt(body.querySelector('[data-field="cf-rank"]').value) || 10;
    const preset = presets[body.querySelector('[data-field="cf-preset"]').value];

    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: 'top10',
        rank,
        percent: isPercent,
        bottom: isBottom,
        dxf: {
            fill: { fgColor: preset.fill },
            font: preset.font ? { color: preset.font } : null
        }
    });
    SN._r();
}

/**
 * 数据条自定义弹窗
 */
async function showDataBarDialog(sqref) {
    const SN = this.SN;
    const colors = ['#638EC6', '#63C384', '#FF555A', '#FFB628', '#8FAADC', '#C6EFCE'];

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.dataBar.fillType')}</span>
                <select class="sn-cf-form-select" data-field="cf-bartype">
                    <option value="gradient">${SN.t('action.condFormat.content.dataBar.gradient')}</option>
                    <option value="solid">${SN.t('action.condFormat.content.dataBar.solid')}</option>
                </select>
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.dataBar.barColor')}</span>
                <input type="color" data-field="cf-barcolor" value="${colors[0]}">
            </div>
            <div class="sn-cf-form-row">
                <label><input type="checkbox" data-field="cf-showvalue" checked> ${SN.t('action.condFormat.content.dataBar.showValue')}</label>
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview" style="position:relative;overflow:hidden;">
                    <div data-field="cf-bar-preview" style="position:absolute;left:0;top:0;height:100%;width:70%;opacity:0.8"></div>
                    <span style="position:relative">75</span>
                </div>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m006.title',
        title: SN.t('action.condFormat.content.modalTitle', { title: SN.t('action.condFormat.content.title.dataBar') }),
        content: html,
        confirmKey: 'action.condFormat.modal.m006.confirm',
        confirmText: SN.t('action.condFormat.modal.m006.confirm'),
        cancelKey: 'action.condFormat.modal.m006.cancel',
        cancelText: SN.t('action.condFormat.modal.m006.cancel')
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const barType = body.querySelector('[data-field="cf-bartype"]').value;
    const barColor = body.querySelector('[data-field="cf-barcolor"]').value;
    const showValue = body.querySelector('[data-field="cf-showvalue"]').checked;

    applyDataBar.call(this, sqref, barType === 'gradient', barColor, showValue);
}

/**
 * 色阶自定义弹窗
 */
async function showColorScaleDialog(sqref) {
    const SN = this.SN;

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.colorScale.type')}</span>
                <select class="sn-cf-form-select" data-field="cf-scaletype">
                    <option value="2">${SN.t('action.condFormat.content.colorScale.twoColors')}</option>
                    <option value="3">${SN.t('action.condFormat.content.colorScale.threeColors')}</option>
                </select>
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.colorScale.minColor')}</span>
                <input type="color" data-field="cf-color-min" value="#F8696B">
            </div>
            <div class="sn-cf-form-row" data-field="cf-mid-row" style="display:none">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.colorScale.midColor')}</span>
                <input type="color" data-field="cf-color-mid" value="#FFEB84">
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.colorScale.maxColor')}</span>
                <input type="color" data-field="cf-color-max" value="#63BE7B">
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview" style="height:24px"></div>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m007.title',
        title: SN.t('action.condFormat.content.modalTitle', { title: SN.t('action.condFormat.content.title.colorScale') }),
        content: html,
        confirmKey: 'action.condFormat.modal.m007.confirm',
        confirmText: SN.t('action.condFormat.modal.m007.confirm'),
        cancelKey: 'action.condFormat.modal.m007.cancel',
        cancelText: SN.t('action.condFormat.modal.m007.cancel')
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const scaleType = body.querySelector('[data-field="cf-scaletype"]').value;
    const colorMin = body.querySelector('[data-field="cf-color-min"]').value;
    const colorMid = body.querySelector('[data-field="cf-color-mid"]').value;
    const colorMax = body.querySelector('[data-field="cf-color-max"]').value;

    const colors = scaleType === '3' ? [colorMin, colorMid, colorMax] : [colorMin, colorMax];
    applyColorScale.call(this, sqref, colors);
}

/**
 * 图标集自定义弹窗
 */
async function showIconSetDialog(sqref) {
    const SN = this.SN;
    const iconSetNames = [
        '3Arrows', '3ArrowsGray', '3Triangles', '4Arrows', '4ArrowsGray', '5Arrows', '5ArrowsGray',
        '3TrafficLights1', '3TrafficLights2', '4TrafficLights', '3Signs', '4RedToBlack',
        '3Symbols', '3Symbols2', '3Flags',
        '3Stars', '4Rating', '5Rating', '5Quarters', '5Boxes'
    ];

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.iconSet.style')}</span>
                <select class="sn-cf-form-select" data-field="cf-iconset" style="width:100%">
                    ${iconSetNames.map(n => `<option value="${n}">${n}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-form-row">
                <label><input type="checkbox" data-field="cf-reverse"> ${SN.t('action.condFormat.content.iconSet.reverse')}</label>
            </div>
            <div class="sn-cf-form-row">
                <label><input type="checkbox" data-field="cf-showvalue" checked> ${SN.t('action.condFormat.content.iconSet.showValue')}</label>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m008.title',
        title: SN.t('action.condFormat.content.modalTitle', { title: SN.t('action.condFormat.content.title.iconSet') }),
        content: html,
        confirmKey: 'action.condFormat.modal.m008.confirm',
        confirmText: SN.t('action.condFormat.modal.m008.confirm'),
        cancelKey: 'action.condFormat.modal.m008.cancel',
        cancelText: SN.t('action.condFormat.modal.m008.cancel')
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const iconSet = body.querySelector('[data-field="cf-iconset"]').value;
    const reverse = body.querySelector('[data-field="cf-reverse"]').checked;
    const showValue = body.querySelector('[data-field="cf-showvalue"]').checked;

    applyIconSet.call(this, sqref, iconSet, reverse, showValue);
}

/**
 * 新建规则弹窗
 */
async function showNewRuleDialog(sqref) {
    const SN = this.SN;

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.newRule.selectType')}</span>
            </div>
            <div class="sn-cf-form-row">
                <select class="sn-cf-form-select" data-field="cf-ruletype" style="width:100%">
                    <option value="cellIs">${SN.t('action.condFormat.content.newRule.cellIs')}</option>
                    <option value="containsText">${SN.t('action.condFormat.content.newRule.containsText')}</option>
                    <option value="top10">${SN.t('action.condFormat.content.newRule.top10')}</option>
                    <option value="aboveAverage">${SN.t('action.condFormat.content.newRule.aboveAverage')}</option>
                    <option value="duplicateValues">${SN.t('action.condFormat.content.newRule.duplicateValues')}</option>
                    <option value="dataBar">${SN.t('action.condFormat.content.newRule.dataBar')}</option>
                    <option value="colorScale">${SN.t('action.condFormat.content.newRule.colorScale')}</option>
                    <option value="iconSet">${SN.t('action.condFormat.content.newRule.iconSet')}</option>
                    <option value="expression">${SN.t('action.condFormat.content.newRule.expression')}</option>
                </select>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m009.title',
        title: SN.t('action.condFormat.modal.m009.title'),
        content: html,
        confirmKey: 'action.condFormat.modal.m009.confirm',
        confirmText: SN.t('action.condFormat.modal.m009.confirm'),
        cancelKey: 'action.condFormat.modal.m009.cancel',
        cancelText: SN.t('action.condFormat.modal.m009.cancel')
    });
    if (!result) throw 'cancel';
    const ruleType = result.bodyEl.querySelector('[data-field="cf-ruletype"]').value;

    // 根据选择打开对应弹窗
    switch (ruleType) {
        case 'cellIs':
            await showCellIsDialog.call(this, sqref, 'greaterThan');
            break;
        case 'containsText':
            await showTextDialog.call(this, sqref, 'containsText');
            break;
        case 'top10':
            await showTop10Dialog.call(this, sqref, 'top');
            break;
        case 'aboveAverage':
            applyAboveAverage.call(this, sqref, 'above');
            break;
        case 'duplicateValues':
            await showDuplicateDialog.call(this, sqref);
            break;
        case 'dataBar':
            await showDataBarDialog.call(this, sqref);
            break;
        case 'colorScale':
            await showColorScaleDialog.call(this, sqref);
            break;
        case 'iconSet':
            await showIconSetDialog.call(this, sqref);
            break;
        case 'expression':
            await showExpressionDialog.call(this, sqref);
            break;
    }
}

/**
 * 公式规则弹窗
 */
async function showExpressionDialog(sqref) {
    const SN = this.SN;
    const presets = getFormatPresets(SN);

    const html = `
        <div class="sn-cf-form">
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.expression.description')}</span>
            </div>
            <div class="sn-cf-form-row">
                <input type="text" class="sn-cf-form-input" data-field="cf-formula" placeholder="=A1>100">
            </div>
            <div class="sn-cf-form-row">
                <span class="sn-cf-form-label">${SN.t('action.condFormat.content.setTo')}</span>
                <select class="sn-cf-form-select" data-field="cf-preset">
                    ${presets.map((preset, i) => `<option value="${i}">${preset.label}</option>`).join('')}
                </select>
            </div>
            <div class="sn-cf-preview-box">
                <span class="sn-cf-preview-label">${SN.t('action.condFormat.content.preview')}</span>
                <div class="sn-cf-preview-cell" data-field="cf-preview">AaBbCc</div>
            </div>
        </div>
    `;

    const result = await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m010.title',
        title: SN.t('action.condFormat.content.modalTitle', { title: SN.t('action.condFormat.content.title.expression') }),
        content: html,
        confirmKey: 'action.condFormat.modal.m010.confirm',
        confirmText: SN.t('action.condFormat.modal.m010.confirm'),
        cancelKey: 'action.condFormat.modal.m010.cancel',
        cancelText: SN.t('action.condFormat.modal.m010.cancel')
    });
    if (!result) throw 'cancel';
    const body = result.bodyEl;

    const formula = body.querySelector('[data-field="cf-formula"]')?.value?.trim();
    const preset = presets[body.querySelector('[data-field="cf-preset"]').value];

    if (!formula) {
        SN.Utils.toast(SN.t('action.condFormat.toast.msg004'));
        throw 'cancel';
    }

    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: 'expression',
        formula,
        dxf: {
            fill: { fgColor: preset.fill },
            font: preset.font ? { color: preset.font } : null
        }
    });
    SN._r();
}

/**
 * 管理规则弹窗
 */
async function showManageDialog() {
    const SN = this.SN;
    const sheet = SN.activeSheet;
    const cf = sheet.CF;
    const rules = cf.getAll();

    if (rules.length === 0) {
        SN.Utils.toast(SN.t('action.condFormat.toast.msg005'));
        return;
    }

    const getRuleDesc = (rule) => {
        const typeMap = {
            cellIs: 'action.condFormat.content.ruleType.cellIs',
            colorScale: 'action.condFormat.content.ruleType.colorScale',
            dataBar: 'action.condFormat.content.ruleType.dataBar',
            iconSet: 'action.condFormat.content.ruleType.iconSet',
            containsText: 'action.condFormat.content.ruleType.containsText',
            duplicateValues: 'action.condFormat.content.ruleType.duplicateValues',
            uniqueValues: 'action.condFormat.content.ruleType.uniqueValues',
            top10: 'action.condFormat.content.ruleType.top10',
            aboveAverage: 'action.condFormat.content.ruleType.aboveAverage',
            timePeriod: 'action.condFormat.content.ruleType.timePeriod',
            expression: 'action.condFormat.content.ruleType.expression'
        };
        const key = typeMap[rule.type];
        return key ? SN.t(key) : rule.type;
    };

    const html = `
        <div class="sn-cf-rule-list">
            ${rules.map((r, i) => `
                <div class="sn-cf-rule-item" data-index="${i}" data-sqref="${r.sqref}" data-priority="${r.priority}">
                    <div class="sn-cf-rule-info">
                        <div class="sn-cf-rule-type">${getRuleDesc(r)}</div>
                        <div class="sn-cf-rule-range">${SN.t('action.condFormat.content.manage.appliesTo', { sqref: r.sqref })}</div>
                    </div>
                    <div class="sn-cf-rule-actions">
                        <button class="sn-cf-rule-btn danger" data-action="delete">${SN.t('action.condFormat.content.manage.delete')}</button>
                    </div>
                </div>
            `).join('')}
        </div>
    `;

    await SN.Utils.modal({
        titleKey: 'action.condFormat.modal.m011.title',
        title: SN.t('action.condFormat.modal.m011.title'),
        content: html,
        maxWidth: '400px',
        confirmKey: 'action.condFormat.modal.m011.confirm',
        confirmText: SN.t('action.condFormat.modal.m011.confirm'),
        onOpen: (bodyEl) => {
            bodyEl.addEventListener('click', (e) => {
                const btn = e.target.closest('[data-action="delete"]');
                if (btn) {
                    const item = btn.closest('.sn-cf-rule-item');
                    const priority = parseInt(item.dataset.priority);
                    const ruleIndex = cf.rules.findIndex(r => r.priority === priority);
                    if (ruleIndex !== -1) {
                        cf.remove(ruleIndex);
                        item.remove();
                        SN._r();
                    }
                }
            });
        }
    });
}

// ==================== 快捷应用方法 ====================

/**
 * 应用高于/低于平均值
 */
function applyAboveAverage(sqref, subType) {
    const SN = this.SN;
    const preset = FORMAT_PRESETS[subType === 'above' ? 2 : 0]; // 绿色或红色

    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: 'aboveAverage',
        aboveAverage: subType === 'above',
        dxf: {
            fill: { fgColor: preset.fill },
            font: preset.font ? { color: preset.font } : null
        }
    });
    SN._r();
}

/**
 * 应用数据条
 */
function applyDataBar(sqref, gradient, color, showValue = true) {
    const SN = this.SN;
    const cf = SN.activeSheet.CF;
    cf.add({
        rangeRef: sqref,
        type: 'dataBar',
        color,
        gradient,
        showValue,
        minCfvo: { type: 'min' },
        maxCfvo: { type: 'max' }
    });
    SN._r();
}

/**
 * 应用色阶
 */
function applyColorScale(sqref, colors) {
    const SN = this.SN;
    const cf = SN.activeSheet.CF;

    const cfvos = colors.length === 2
        ? [{ type: 'min' }, { type: 'max' }]
        : [{ type: 'min' }, { type: 'percentile', val: 50 }, { type: 'max' }];

    cf.add({
        rangeRef: sqref,
        type: 'colorScale',
        cfvos,
        colors
    });
    SN._r();
}

/**
 * 应用图标集
 */
function applyIconSet(sqref, iconSet, reverse = false, showValue = true) {
    const SN = this.SN;
    const cf = SN.activeSheet.CF;

    // 根据图标集名称确定数量
    const count = iconSet.startsWith('5') ? 5 : iconSet.startsWith('4') ? 4 : 3;
    const cfvos = [];
    for (let i = 0; i < count; i++) {
        cfvos.push({ type: 'percent', val: Math.round((i / count) * 100) });
    }

    cf.add({
        rangeRef: sqref,
        type: 'iconSet',
        iconSet,
        cfvos,
        reverse,
        showValue
    });
    SN._r();
}

/**
 * 清除选区规则
 */
function clearSelectionRules(sqref) {
    const SN = this.SN;
    const cf = SN.activeSheet.CF;
    cf.removeAll(sqref);
    SN._r();
    SN.Utils.toast(SN.t('action.condFormat.toast.msg006'));
}

/**
 * 清除工作表所有规则
 */
function clearSheetRules() {
    const SN = this.SN;
    const cf = SN.activeSheet.CF;
    const rules = cf.getAll();
    const sqrefs = [...new Set(rules.map(r => r.sqref))];
    sqrefs.forEach(sq => cf.removeAll(sq));
    SN._r();
    SN.Utils.toast(SN.t('action.condFormat.toast.msg007'));
}

// ==================== 工具函数 ====================

/**
 * 解析输入值
 */
function parseValue(str) {
    if (!str) return null;
    const num = parseFloat(str);
    return isNaN(num) ? str : num;
}
