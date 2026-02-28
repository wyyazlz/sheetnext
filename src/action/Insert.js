// 插入操作模块
import { fileToBase64 } from './helpers.js';
import {
    CHART_CATEGORIES,
    getChartDisplayLabel,
    getChartTypeKeyByLabel,
    getLocalizedChartCategories
} from '../components/ChartConfig.js';
import { getChartSvg } from '../assets/chartSvgs.js';

const CHART_TYPE_ALIASES = {
    column: 'bar_clustered',
    bar: 'barh_clustered',
    line: 'line_basic',
    pie: 'pie_basic',
    doughnut: 'pie_doughnut',
    area: 'area_basic',
    scatter: 'scatter_basic',
    radar: 'radar_basic',
    histogram: 'histogram_basic'
};

const SUPPORTED_CHART_BASES = new Set([
    'bar',
    'histogram',
    'line',
    'area',
    'pie',
    'scatter',
    'radar',
    'stock',
    'sunburst',
    'treemap',
    'funnel',
    'waterfall',
    'combo'
]);

const SPARKLINE_TYPE_LABELS = {
    line: 'action.insert.content.sparkline.type.line',
    column: 'action.insert.content.sparkline.type.column',
    win_loss: 'action.insert.content.sparkline.type.winLoss'
};

const TEXT_BOX_PRESETS = {
    horizontal: { width: 160, height: 90 },
    vertical: { width: 90, height: 160 }
};

function resolveChartMeta(SN, chartType) {
    if (!chartType) return null;
    const rawType = String(chartType).trim();
    if (!rawType) return null;

    const lowered = rawType.toLowerCase();
    const localizedKey = getChartTypeKeyByLabel(SN, rawType);
    const normalized = CHART_TYPE_ALIASES[lowered] || localizedKey || rawType;

    let baseKey = normalized;
    let variant = '';
    if (normalized.includes('_')) {
        const parts = normalized.split('_');
        baseKey = parts.shift();
        variant = parts.join('_');
    }

    if (!variant) {
        const defaultVariants = {
            bar: 'clustered',
            barh: 'clustered',
            histogram: 'basic',
            line: 'basic',
            area: 'basic',
            pie: 'basic',
            scatter: 'basic',
            radar: 'basic'
        };
        variant = defaultVariants[baseKey] || '';
    }

    let orientation = 'vertical';
    let base = baseKey;
    if (baseKey === 'barh') {
        base = 'bar';
        orientation = 'horizontal';
    }

    const stacked = variant.includes('stacked') || variant.includes('percent');
    const percent = variant.includes('percent');
    const markers = variant.includes('marker');
    const smooth = variant.includes('smooth');
    const doughnut = variant.includes('doughnut');
    const bubble = variant.includes('bubble');
    const radarFilled = variant.includes('filled');
    const lineLike = variant.includes('line') || smooth;

    const category = CHART_CATEGORIES.find((item) => item.key === baseKey);
    const categoryLabel = category ? SN.t(category.labelKey) : rawType;
    const label = getChartDisplayLabel(SN, normalized, categoryLabel || rawType);

    return {
        key: normalized,
        base,
        baseKey,
        variant,
        label,
        orientation,
        stacked,
        percent,
        markers,
        smooth,
        doughnut,
        radarFilled,
        bubble,
        lineLike,
        supported: SUPPORTED_CHART_BASES.has(base)
    };
}

/**
 * 插入批注
 */
export async function insertComment() {
    const sheet = this.SN.activeSheet;
    const activeCell = sheet.activeCell;
    const cellRef = this.SN.Utils.cellNumToStr(activeCell);

    // 检查是否已有批注
    const existingComment = sheet.Comment.get(cellRef);
    const modalTitleKey = existingComment
        ? 'action.insert.modal.m001.titleEdit'
        : 'action.insert.modal.m001.titleInsert';

    const result = await this.SN.Utils.modal({
        titleKey: modalTitleKey,
        title: this.SN.t(modalTitleKey),
        maxWidth:'400px',
        content: `
            <div style="margin-bottom: 10px;">
                <label style="display: block; margin-bottom: 5px;">Cell: ${cellRef}</label>
            </div>
            <div style="margin-bottom: 10px;">
                <label style="display: block; margin-bottom: 5px;">Note:</label>
                <textarea data-field="comment-text" style="width: 100%;border:1px solid #ccc; height: 80px; resize: none; padding: 8px; box-sizing: border-box;">${existingComment?.text || ''}</textarea>
            </div>
        `,
        confirmKey: 'action.insert.modal.m001.confirm', confirmText: 'OK',
        cancelKey: 'action.insert.modal.m001.cancel', cancelText: 'Cancel'
    });
    if (!result) return;

    const textarea = result.bodyEl.querySelector('[data-field="comment-text"]');
    const text = textarea?.value?.trim();

    if (!text) {
        // 如果内容为空且存在批注，则删除批注
        if (existingComment) {
            sheet.Comment.remove(cellRef);
            this.SN._r();
        }
        return;
    }

    if (existingComment) {
        // 更新已有批注
        existingComment.text = text;
    } else {
        // 添加新批注
        sheet.Comment.add({ cellRef, text });
    }

    this.SN._r();
}

/**
 * 插入区域截图 - 将当前选中区域截图并作为图片插入
 */
export async function insertAreaScreenshot() {
    const sheet = this.SN.activeSheet;
    const area = sheet.activeAreas[sheet.activeAreas.length - 1];
    if (!area) return;

    // 计算选中区域的像素尺寸
    let areaWidth = 0;
    let areaHeight = 0;
    for (let c = area.s.c; c <= area.e.c; c++) {
        const col = sheet.getCol(c);
        if (!col.hidden) areaWidth += col.width;
    }
    for (let r = area.s.r; r <= area.e.r; r++) {
        const row = sheet.getRow(r);
        if (!row.hidden) areaHeight += row.height;
    }
    if (areaWidth <= 0 || areaHeight <= 0) return this.SN.Utils.toast(this.SN.t('action.insert.toast.msg001'));

    const headerW = sheet.indexWidth;
    const headerH = sheet.headHeight;
    const cellRef = this.SN.Utils.cellNumToStr(area.s);

    const dataURL = await this.SN.Canvas.captureScreenshot(cellRef, 'topleft', {
        width: areaWidth + headerW,
        height: areaHeight + headerH,
        compress: false,
        hideSelectionFrame: true
    });

    // 裁剪掉行列头区域（captureScreenshot 内部 exportScale=2）
    const scale = 2;
    const img = new Image();
    img.src = dataURL;
    await new Promise(resolve => { img.onload = resolve; });

    const cropCanvas = document.createElement('canvas');
    cropCanvas.width = Math.round(areaWidth * scale);
    cropCanvas.height = Math.round(areaHeight * scale);
    const ctx = cropCanvas.getContext('2d');
    ctx.drawImage(img,
        Math.round(headerW * scale), Math.round(headerH * scale),
        cropCanvas.width, cropCanvas.height,
        0, 0, cropCanvas.width, cropCanvas.height
    );

    // 指定 width/height 避免 _initImage 按 naturalWidth 自动缩放
    sheet.Drawing.addImage(cropCanvas.toDataURL('image/png'), {
        width: areaWidth,
        height: areaHeight
    });
    this.SN._r();
}

export async function insertImages() {
    if (!window.showOpenFilePicker) {
        this.SN.Utils.toast(this.SN.t('action.insert.content.imagePickerNotSupported'));
        return;
    }

    // 打开文件选择器，允许选择多张图片
    const fileHandles = await window.showOpenFilePicker({
        types: [
            {
                description: this.SN.t('action.insert.content.imageFileDescription'),
                accept: {
                    'image/*': ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg']
                }
            }
        ],
        multiple: true  // 允许选择多个文件
    });

    const sheet = this.SN.activeSheet;
    let offsetX = 0;
    let offsetY = 0;

    // 处理每个选中的图片文件
    for (let i = 0; i < fileHandles.length; i++) {
        const fileHandle = fileHandles[i];
        const file = await fileHandle.getFile();

        // 将文件转换为 Base64
        const base64 = await fileToBase64(file);

        // 插入图片到表格中（尺寸由 Drawing 自动处理）
        sheet.Drawing.addImage(base64, {
            offsetX: offsetX,
            offsetY: offsetY
        });

        // 为下一张图片设置水平和垂直偏移，避免重叠
        offsetX += 20;  // 水平偏移 20像素
        offsetY += 20;  // 垂直偏移 20像素
    }

    this.SN._r();
}

function normalizeTextBoxDirection(direction) {
    return direction === 'vertical' ? 'vertical' : 'horizontal';
}

/**
 * 插入文本框（横排/竖排）
 * @param {'horizontal'|'vertical'} [direction='horizontal'] - 文本方向
 * @returns {boolean}
 */
export function insertTextBox(direction = 'horizontal') {
    const sheet = this.SN.activeSheet;
    const textDirection = normalizeTextBoxDirection(direction);
    const preset = TEXT_BOX_PRESETS[textDirection] || TEXT_BOX_PRESETS.horizontal;
    const textAlign = textDirection === 'vertical' ? 'r' : 'l';
    const activeCell = sheet.activeCell || { r: 0, c: 0 };

    const drawing = sheet.Drawing.addShape('rect', {
        isTextBox: true,
        startCell: { r: activeCell.r, c: activeCell.c },
        width: preset.width,
        height: preset.height,
        shapeText: this.SN.t('action.insert.content.textBox.defaultText'),
        shapeStyle: {
            fill: '#FFFFFF',
            stroke: '#7F7F7F',
            strokeWidth: 0.75,
            textColor: '#000000',
            fontSize: 11,
            textAlign,
            textDirection,
            startArrow: 'none',
            endArrow: 'none'
        }
    });

    if (!drawing) return false;
    this.SN._r();
    return true;
}

function buildChartModalHTML(categories) {
    const tabs = categories.map((cat) => (
        `<span class="sn-chart-tab" data-category="${cat.key}">${cat.text}</span>`
    )).join('');

    const panels = categories.map((cat) => {
        const items = cat.items.map((item) => (
            `<div class="sn-chart-item" data-category="${cat.key}" data-item="${item.key}" title="${item.text}">
                <div class="sn-chart-preview">${getChartSvg(item.key)}</div>
                <div class="sn-chart-item-label">${item.text}</div>
            </div>`
        )).join('');
        return `<div class="sn-chart-panel" data-category="${cat.key}">${items}</div>`;
    }).join('');

    return `
        <div class="sn-chart-modal">
            <div class="sn-chart-tabs">${tabs}</div>
            <div class="sn-chart-panels">${panels}</div>
        </div>
    `;
}

function initChartModalInteraction(bodyEl, initialCategory, initialItem) {
    const modalEl = bodyEl.querySelector('.sn-chart-modal');
    if (!modalEl) return;

    const tabs = Array.from(modalEl.querySelectorAll('.sn-chart-tab'));
    const panels = Array.from(modalEl.querySelectorAll('.sn-chart-panel'));

    const selectItem = (itemEl) => {
        if (!itemEl) return;
        modalEl.querySelectorAll('.sn-chart-item.selected').forEach((el) => el.classList.remove('selected'));
        itemEl.classList.add('selected');
        modalEl.dataset.selectedCategory = itemEl.dataset.category || '';
        modalEl.dataset.selectedItem = itemEl.dataset.item || '';
    };

    const setActiveCategory = (key) => {
        const targetKey = key || CHART_CATEGORIES[0]?.key;
        tabs.forEach((tab) => tab.classList.toggle('active', tab.dataset.category === targetKey));
        panels.forEach((panel) => panel.classList.toggle('active', panel.dataset.category === targetKey));
        modalEl.dataset.selectedCategory = targetKey;

        const activePanel = modalEl.querySelector(`.sn-chart-panel[data-category="${targetKey}"]`);
        if (!activePanel?.querySelector('.sn-chart-item.selected')) {
            selectItem(activePanel?.querySelector('.sn-chart-item'));
        }
    };

    setActiveCategory(initialCategory?.key);

    if (initialItem) {
        const initialItemEl = modalEl.querySelector(`.sn-chart-item[data-item="${initialItem}"]`);
        if (initialItemEl) {
            setActiveCategory(initialItemEl.dataset.category);
            selectItem(initialItemEl);
        }
    }

    modalEl.addEventListener('click', (event) => {
        const tabEl = event.target.closest('.sn-chart-tab');
        if (tabEl) {
            setActiveCategory(tabEl.dataset.category);
            return;
        }
        const itemEl = event.target.closest('.sn-chart-item');
        if (itemEl) selectItem(itemEl);
    });
}

export async function openChartModal(categoryKey, itemKey) {
    const categories = getLocalizedChartCategories(this.SN);
    const defaultCategory = categories.find((cat) => cat.key === categoryKey) || categories[0];

    try {
        await this.SN.Utils.modal({
            titleKey: 'chart.modal.title',
            title: 'Insert Chart',
            content: buildChartModalHTML(categories),
            maxWidth: '580px',
            confirmKey: 'chart.modal.confirm',
            confirmText: 'Insert',
            cancelKey: 'common.cancel',
            cancelText: 'Cancel',
            onOpen: (bodyEl) => {
                initChartModalInteraction(bodyEl, defaultCategory, itemKey);
            },
            onConfirm: (bodyEl) => {
                const modalEl = bodyEl.querySelector('.sn-chart-modal');
                if (!modalEl) return;
                const selectedItem = modalEl.dataset.selectedItem || '';
                if (selectedItem) {
                    this.insertChart(selectedItem);
                }
            }
        });
    } catch (e) {
        // 用户取消
    }
}

export function insertChart(chartType) {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas;

    if (!areas || areas.length === 0) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.selectRange'));
    }

    // 获取选中区域
    const area = areas[areas.length - 1];

    // 读取选中区域数据
    let data = [];
    for (let r = area.s.r; r <= area.e.r; r++) {
        const row = [];
        for (let c = area.s.c; c <= area.e.c; c++) {
            const cell = sheet.getCell(r, c);
            row.push(cell.showVal || '');
        }
        data.push(row);
    }

    if (data.length === 0) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.rangeEmpty'));
    }

    // 检测数据方向：数据行数 > 数据列数时，分类在第一列，需要转置
    const dataRows = data.length - 1;
    const dataCols = (data[0]?.length || 1) - 1;
    const transposed = dataRows > dataCols;

    if (transposed) {
        const cols = data[0].length;
        const rows = data.length;
        const newData = [];
        for (let c = 0; c < cols; c++) {
            const row = [];
            for (let r = 0; r < rows; r++) {
                row.push(data[r][c] ?? '');
            }
            newData.push(row);
        }
        data = newData;
    }

    // 解析数据：第一行作为类别/标签，第一列作为系列名称
    const categories = data[0].slice(1);  // 第一行，跳过第一列
    const meta = resolveChartMeta(this.SN, chartType || this.SN.chartSelection?.item);
    if (!meta) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.noType'));
    }
    if (!meta.supported) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.unsupportedType', { label: meta.label }));
    }

    const series = [];

    for (let i = 1; i < data.length; i++) {
        const seriesName = data[i][0];
        const seriesData = data[i].slice(1).map(v => {
            const num = parseFloat(v);
            return isNaN(num) ? 0 : num;
        });
        series.push({
            name: seriesName,
            data: seriesData
        });
    }

    // 生成 ECharts 配置
    const option = generateChartOption(this.SN, meta, categories, series);
    if (!option) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.unsupportedType', { label: meta.label }));
    }
    if (meta.percent) option.__percent = true;
    const excelRef = buildChartExcelRefs(sheet, area, transposed);
    option.__excelRef = excelRef;

    // 标准图表类型：用单元格引用替换原始值，实现数据联动
    const specialTypes = ['sunburst', 'treemap', 'funnel', 'waterfall', 'histogram', 'stock', 'combo', 'pie', 'radar', 'scatter'];
    if (!specialTypes.includes(meta.base)) {
        if (option.xAxis?.type === 'category') option.xAxis.data = excelRef.categories;
        if (option.yAxis?.type === 'category') option.yAxis.data = excelRef.categories;
        option.series?.forEach((s, i) => {
            if (excelRef.series[i]) {
                s.name = excelRef.series[i].name;
                s.data = excelRef.series[i].data;
            }
        });
        if (option.legend?.data) {
            option.legend.data = excelRef.series.map(s => s.name);
        }
    }

    if (['sunburst', 'treemap', 'funnel', 'waterfall', 'histogram'].includes(meta.base)) {
        option.__chartType = meta.base;
        option.__excelHidden = excelRef.series.map((_, index) => index > 0);
        if (meta.base === 'waterfall') {
            option.__waterfall = {
                categories: excelRef.categories,
                values: excelRef.series?.[0]?.data || '',
                seriesName: series[0]?.name || meta.label
            };
        } else if (meta.base === 'histogram') {
            option.__histogram = { intervalClosed: 'r' };
        }
    }

    // 插入图表
    sheet.Drawing.addChart(option);

    this.SN._r();
    return true;
}

function normalizeSparklineKey(type) {
    const raw = String(type || '').trim().toLowerCase();
    if (!raw) return 'line';
    if (raw === 'stacked' || raw === 'winloss' || raw === 'win-loss' || raw === 'win_loss') return 'win_loss';
    if (raw === 'column' || raw === 'line') return raw;
    return 'line';
}

function normalizeSparklineType(type) {
    const key = normalizeSparklineKey(type);
    return key === 'win_loss' ? 'stacked' : key;
}

function resolveSparklineRange(SN, defaultSheet, rangeInput) {
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
    } catch (e) {
        throw SN.t('action.insert.content.sparkline.rangeInvalid');
    }
    if (!rangeObj) throw SN.t('action.insert.content.sparkline.rangeInvalid');

    const formula = buildRangeRef(sourceSheet, rangeObj.s.r, rangeObj.s.c, rangeObj.e.r, rangeObj.e.c);
    return { sourceSheet, rangeObj, formula };
}

function resolveSparklineLocation(SN, defaultSheet, locationInput) {
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
    const formula = buildRangeRef(sheet, area.s.r, area.s.c, area.e.r, area.e.c);
    const targetCell = sheet.activeCell;
    const normalizedType = normalizeSparklineType(type);

    sheet.Sparkline.add({
        r: targetCell.r,
        c: targetCell.c,
        formula,
        type: normalizedType
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
    const defaultTypeKey = normalizeSparklineKey(type);

    const typeOptions = Object.entries(SPARKLINE_TYPE_LABELS).map(([key, labelKey]) => (
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

                const { formula } = resolveSparklineRange(SN, sheet, rangeInput);
                const { targetSheet, cellObj } = resolveSparklineLocation(SN, sheet, locationInput);

                if (!targetSheet.Sparkline) throw SN.t('action.insert.content.sparkline.moduleNotReady');

                targetSheet.Sparkline.add({
                    r: cellObj.r,
                    c: cellObj.c,
                    formula,
                    type: normalizeSparklineType(typeValue)
                });

                SN.activeSheet = targetSheet;
                SN._r();
            }
        });
    } catch (e) {
        // 用户取消
    }
}

/**
 * 插入/编辑超链接
 */
export async function insertHyperlink() {
    const sheet = this.SN.activeSheet;
    const activeCell = sheet.activeCell;
    const cell = sheet.getCell(activeCell.r, activeCell.c);
    const existingLink = cell.hyperlink;
    const currentText = cell.editVal || '';

    // 获取所有工作表名称
    const sheetNames = this.SN.sheets.map(s => s.name);
    const sheetOptions = sheetNames.map(n => `<option value="${n}">${n}</option>`).join('');

    // 解析已有的 location
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
    const SN = this.SN;
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
            confirmKey: 'action.insert.modal.m004.confirm', confirmText: 'OK',
            maxWidth:'400px',
            cancelKey: modalCancelKey,
            cancelText: SN.t(modalCancelKey),
            onOpen(bodyEl) {
                // 设置初始工作表值
                bodyEl.querySelector('[data-field="link-sheet"]').value = existingSheet;
                // 链接类型切换
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
    } catch (e) {
        // 用户取消
    }
}

function toNumber(value, fallback) {
    const num = Number(value);
    return Number.isFinite(num) ? num : fallback;
}

function normalizeSheetName(name) {
    return name.includes(' ') ? `'${name}'` : name;
}

function buildCellRef(sheet, r, c) {
    const sheetName = normalizeSheetName(sheet.name);
    const cellRef = sheet.Utils.cellNumToStr({ r, c });
    return `${sheetName}!${cellRef}`;
}

function buildRangeRef(sheet, sr, sc, er, ec) {
    const sheetName = normalizeSheetName(sheet.name);
    const start = sheet.Utils.cellNumToStr({ r: sr, c: sc });
    const end = sheet.Utils.cellNumToStr({ r: er, c: ec });
    return `${sheetName}!${start}:${end}`;
}

function buildChartExcelRefs(sheet, area, transposed) {
    if (transposed) {
        const categories = buildRangeRef(sheet, area.s.r + 1, area.s.c, area.e.r, area.s.c);
        const series = [];
        for (let c = area.s.c + 1; c <= area.e.c; c++) {
            series.push({
                name: buildCellRef(sheet, area.s.r, c),
                data: buildRangeRef(sheet, area.s.r + 1, c, area.e.r, c)
            });
        }
        return { categories, series };
    }
    const categories = buildRangeRef(sheet, area.s.r, area.s.c + 1, area.s.r, area.e.c);
    const series = [];
    for (let r = area.s.r + 1; r <= area.e.r; r++) {
        series.push({
            name: buildCellRef(sheet, r, area.s.c),
            data: buildRangeRef(sheet, r, area.s.c + 1, r, area.e.c)
        });
    }
    return { categories, series };
}

function normalizeSeriesToPercent(series) {
    if (!series.length) return [];
    const length = series[0].data.length;
    const totals = new Array(length).fill(0);

    series.forEach((s) => {
        s.data.forEach((val, idx) => {
            totals[idx] += val;
        });
    });

    return series.map((s) => ({
        name: s.name,
        data: s.data.map((val, idx) => {
            const total = totals[idx];
            if (!total) return 0;
            return Math.round((val / total) * 10000) / 100;
        })
    }));
}

function buildScatterSeries(meta, categories, series) {
    const xValues = categories.map((val, idx) => toNumber(val, idx + 1));
    const seriesType = meta.lineLike ? 'line' : 'scatter';
    const showSymbol = meta.lineLike ? meta.markers : true;

    const buildPairs = (data) => data.map((val, idx) => {
        const xVal = xValues[idx] ?? idx + 1;
        return [xVal, val];
    });

    const buildBubbleData = (data) => data.map((val, idx) => {
        const xVal = xValues[idx] ?? idx + 1;
        if (Array.isArray(val)) {
            const yVal = toNumber(val[1] ?? val[0], 0);
            const sizeVal = toNumber(val.length > 2 ? val[2] : yVal, yVal);
            return [xVal, yVal, sizeVal];
        }
        const yVal = toNumber(val, 0);
        return [xVal, yVal, yVal];
    });

    const buildBubbleSymbolSizer = (data) => {
        const sizes = data.map(item => toNumber(item[2] ?? item[1] ?? item[0], 0));
        const min = Math.min(...sizes);
        const max = Math.max(...sizes);
        const minSize = 6;
        const maxSize = 30;
        if (!Number.isFinite(min) || !Number.isFinite(max) || max === min) {
            const fixed = (minSize + maxSize) / 2;
            return {
                min,
                max,
                symbolSize: () => fixed
            };
        }
        return {
            min,
            max,
            symbolSize: (val) => {
                const raw = Array.isArray(val) ? (val[2] ?? val[1] ?? val[0]) : val;
                const num = toNumber(raw, min);
                return minSize + (maxSize - minSize) * ((num - min) / (max - min));
            }
        };
    };

    return series.map((s) => {
        if (meta.bubble && seriesType === 'scatter') {
            const data = buildBubbleData(s.data);
            const bubbleInfo = buildBubbleSymbolSizer(data);
            return {
                name: s.name,
                type: 'scatter',
                data,
                smooth: false,
                showSymbol: true,
                symbolSize: bubbleInfo.symbolSize,
                __bubble: true,
                __bubbleMin: bubbleInfo.min,
                __bubbleMax: bubbleInfo.max
            };
        }
        return {
            name: s.name,
            type: seriesType,
            data: buildPairs(s.data),
            smooth: meta.smooth,
            showSymbol
        };
    });
}

function extractSeriesValues(seriesData, fallback = 0) {
    if (!Array.isArray(seriesData)) return [];
    return seriesData.map((item, idx) => {
        if (typeof item === 'number') return item;
        if (typeof item === 'string') return toNumber(item, fallback);
        if (Array.isArray(item)) {
            const value = item.length > 1 ? item[1] : item[0];
            return toNumber(value, fallback);
        }
        if (item && typeof item === 'object') {
            if (Object.prototype.hasOwnProperty.call(item, 'value')) {
                return toNumber(item.value, fallback);
            }
        }
        return fallback;
    });
}

function buildNameValueItems(categories, series, seriesIndex = 0, SN = null) {
    const values = extractSeriesValues(series[seriesIndex]?.data);
    return (categories || []).map((name, idx) => ({
        name: String(name ?? (SN ? SN.t('action.insert.content.chart.defaultItem', { index: idx + 1 }) : `Item ${idx + 1}`)),
        value: values[idx] ?? 0
    }));
}

function findSeriesByKeywords(series, keywords) {
    const kw = keywords.map(k => String(k).toLowerCase());
    return series.find((s) => {
        const name = String(s?.name ?? '').toLowerCase();
        return kw.some(k => name.includes(k));
    }) || null;
}

function buildStockOption(SN, meta, categories, series) {
    const openSeries = findSeriesByKeywords(series, ['开盘', 'open']);
    const closeSeries = findSeriesByKeywords(series, ['收盘', 'close']);
    const highSeries = findSeriesByKeywords(series, ['最高', 'high']);
    const lowSeries = findSeriesByKeywords(series, ['最低', 'low']);
    const volumeSeries = findSeriesByKeywords(series, ['成交量', 'volume', 'vol']);

    const remaining = series.filter(s => ![openSeries, closeSeries, highSeries, lowSeries, volumeSeries].includes(s));
    const [fallbackA, fallbackB, fallbackC, fallbackD] = remaining;

    const open = openSeries?.data ?? fallbackA?.data ?? [];
    const high = highSeries?.data ?? fallbackB?.data ?? [];
    const low = lowSeries?.data ?? fallbackC?.data ?? [];
    const close = closeSeries?.data ?? fallbackD?.data ?? open;
    const volume = volumeSeries?.data ?? [];

    const candleData = (categories || []).map((_, idx) => ([
        toNumber(open[idx], 0),
        toNumber(close[idx], 0),
        toNumber(low[idx], 0),
        toNumber(high[idx], 0)
    ]));

    const option = {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { data: [] },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: [{
            name: meta.label,
            type: 'candlestick',
            data: candleData
        }]
    };

    if (volume && volume.length) {
        option.series.push({
            name: SN.t('action.insert.content.chart.volume'),
            type: 'bar',
            data: extractSeriesValues(volume)
        });
    }

    return option;
}

function buildFunnelOption(SN, meta, categories, series) {
    const data = buildNameValueItems(categories, series, 0, SN);
    return {
        title: { text: meta.label },
        tooltip: { trigger: 'item' },
        legend: { show: false },
        series: [{
            name: meta.label,
            type: 'funnel',
            sort: 'descending',
            data
        }]
    };
}

function buildTreeOption(meta, categories, series, type) {
    const data = buildNameValueItems(categories, series, 0);
    const legend = type === 'treemap'
        ? { top: 40, left: 'center', data: categories }
        : { show: false };
    return {
        title: { text: meta.label },
        tooltip: { trigger: 'item' },
        legend,
        series: [{
            name: meta.label,
            type,
            data
        }]
    };
}

function buildWaterfallOption(SN, meta, categories, series) {
    const values = extractSeriesValues(series[0]?.data);
    let cumulative = 0;
    const assist = [];
    const actual = values.map((val) => {
        const start = cumulative;
        cumulative += val;
        assist.push(start);
        return val;
    });
    const seriesName = series[0]?.name || meta.label;
    const colorFn = (params) => {
        const val = actual[params.dataIndex] ?? 0;
        return val >= 0 ? '#70AD47' : '#ED7D31';
    };

    return {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { top: 40, left: 'center', data: [seriesName] },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: [
            {
                name: SN.t('action.insert.content.chart.assist'),
                type: 'bar',
                stack: 'total',
                itemStyle: { color: 'transparent' },
                data: assist
            },
            {
                name: seriesName,
                type: 'bar',
                stack: 'total',
                itemStyle: { color: colorFn },
                data: actual
            }
        ]
    };
}

function buildHistogramOption(meta, categories, series) {
    const values = extractSeriesValues(series[0]?.data);
    const seriesName = series[0]?.name || meta.label;
    return {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { show: false },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: [{
            name: seriesName,
            type: 'bar',
            data: values,
            barGap: '0%',
            barCategoryGap: '0%'
        }]
    };
}

function buildComboOption(meta, categories, series) {
    const option = {
        title: { text: meta.label },
        tooltip: { trigger: 'axis' },
        legend: { data: series.map(s => s.name) },
        xAxis: { type: 'category', data: categories },
        yAxis: { type: 'value' },
        series: []
    };

    if (!series.length) return option;
    const barSeries = series[0];
    const lineSeries = series[1];

    if (barSeries) {
        option.series.push({
            name: barSeries.name,
            type: 'bar',
            data: barSeries.data
        });
    }
    if (lineSeries) {
        const lineConfig = {
            name: lineSeries.name,
            type: 'line',
            data: lineSeries.data
        };
        if (meta.variant.includes('area')) {
            lineConfig.areaStyle = {};
        }
        option.series.push(lineConfig);
    } else if (meta.variant.includes('area')) {
        option.series[0].type = 'line';
        option.series[0].areaStyle = {};
    }

    return option;
}

function generateChartOption(SN, meta, categories, series) {
    if (!meta || !meta.supported) return null;

    const title = meta.label || meta.key;
    const option = {
        title: { text: title },
        tooltip: { trigger: meta.base === 'pie' ? 'item' : 'axis' },
        legend: {
            data: meta.base === 'pie' ? categories : series.map(s => s.name)
        }
    };

    if (meta.base === 'stock') {
        return buildStockOption(SN, meta, categories, series);
    }

    if (meta.base === 'sunburst') {
        return buildTreeOption(meta, categories, series, 'sunburst');
    }

    if (meta.base === 'treemap') {
        return buildTreeOption(meta, categories, series, 'treemap');
    }

    if (meta.base === 'funnel') {
        return buildFunnelOption(SN, meta, categories, series);
    }

    if (meta.base === 'waterfall') {
        return buildWaterfallOption(SN, meta, categories, series);
    }

    if (meta.base === 'histogram') {
        return buildHistogramOption(meta, categories, series);
    }

    if (meta.base === 'combo') {
        return buildComboOption(meta, categories, series);
    }

    if (meta.base === 'pie') {
        option.legend.data = categories;
        const pieData = categories.map((name, idx) => ({
            name: name,
            value: series[0]?.data[idx] || 0
        }));

        option.series = [{
            name: series[0]?.name || SN.t('action.insert.content.chart.defaultSeries'),
            type: 'pie',
            radius: meta.doughnut ? ['40%', '70%'] : '60%',
            data: pieData,
            emphasis: {
                itemStyle: {
                    shadowBlur: 10,
                    shadowOffsetX: 0,
                    shadowColor: 'rgba(0, 0, 0, 0.5)'
                }
            },
            label: {
                show: false
            }
        }];
        return option;
    }

    if (meta.base === 'radar') {
        option.radar = {
            indicator: categories.map(name => ({
                name
            }))
        };
        option.series = series.map(s => ({
            name: s.name,
            type: 'radar',
            data: [{
                value: s.data,
                name: s.name
            }],
            symbolSize: meta.markers ? 6 : 0
        }));
        if (meta.radarFilled) {
            option.series = option.series.map(s => ({ ...s, areaStyle: {} }));
        }
        return option;
    }

    if (meta.base === 'scatter') {
        option.xAxis = { type: 'value' };
        option.yAxis = { type: 'value' };
        option.series = buildScatterSeries(meta, categories, series);
        option.tooltip.trigger = meta.lineLike ? 'axis' : 'item';
        return option;
    }

    const usePercent = meta.percent;
    const normalizedSeries = usePercent ? normalizeSeriesToPercent(series) : series;
    const valueAxis = {
        type: 'value',
        max: usePercent ? 100 : undefined,
        axisLabel: usePercent ? { formatter: '{value}%' } : undefined
    };

    if (meta.orientation === 'horizontal') {
        option.xAxis = valueAxis;
        option.yAxis = { type: 'category', data: categories };
    } else {
        option.xAxis = { type: 'category', data: categories };
        option.yAxis = valueAxis;
    }

    option.series = normalizedSeries.map(s => {
        const seriesConfig = {
            name: s.name,
            type: meta.base === 'bar' ? 'bar' : 'line',
            data: s.data
        };

        if (meta.stacked) seriesConfig.stack = 'total';
        if (meta.base === 'area') seriesConfig.areaStyle = {};
        if (meta.base !== 'bar') {
            seriesConfig.showSymbol = meta.markers;
            seriesConfig.smooth = meta.smooth;
        }

        return seriesConfig;
    });

    return option;
}
