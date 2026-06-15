// 插入操作模块
import * as echarts from 'echarts/core';
import {
    BarChart,
    CandlestickChart,
    FunnelChart,
    LineChart,
    PieChart,
    RadarChart,
    ScatterChart,
    SunburstChart,
    TreemapChart
} from 'echarts/charts';
import {
    GridComponent,
    LegendComponent,
    RadarComponent,
    TitleComponent,
    TooltipComponent
} from 'echarts/components';
import { CanvasRenderer } from 'echarts/renderers';
import { fileToBase64 } from './helpers.js';
import { getLocalizedChartCategories } from '../components/ChartConfig.js';
import {
    buildChartModalHTML,
    getSelectedChartItem,
    initChartModalInteraction
} from '../components/ChartModal.js';
import { drawShapeToCanvas } from '../core/Canvas/ShapeRenderer.js';
import { createChartOptionFromRange, resolveChartMeta } from '../core/Drawing/chartOption.js';

echarts.use([
    TitleComponent,
    TooltipComponent,
    LegendComponent,
    GridComponent,
    RadarComponent,
    BarChart,
    LineChart,
    PieChart,
    ScatterChart,
    RadarChart,
    CandlestickChart,
    SunburstChart,
    TreemapChart,
    FunnelChart,
    CanvasRenderer
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

function _renderRotatedCanvasToDataURL(sourceCanvas, logicalW, logicalH, rotation, pixelRatio = 1) {
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(logicalW * pixelRatio));
    canvas.height = Math.max(1, Math.round(logicalH * pixelRatio));
    const ctx = canvas.getContext('2d');
    ctx.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
    ctx.translate(logicalW / 2, logicalH / 2);
    ctx.rotate(rotation * Math.PI / 180);
    ctx.drawImage(sourceCanvas, -logicalW / 2, -logicalH / 2, logicalW, logicalH);
    return canvas.toDataURL('image/png');
}

async function _renderChartDrawingToDataURL(drawing, pixelRatio = 2) {
    const logicalW = Math.max(1, Number(drawing.width) || 1);
    const logicalH = Math.max(1, Number(drawing.height) || 1);
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(logicalW * pixelRatio));
    canvas.height = Math.max(1, Math.round(logicalH * pixelRatio));
    const chart = echarts.init(canvas, null, {
        renderer: 'canvas',
        width: logicalW,
        height: logicalH,
        devicePixelRatio: pixelRatio
    });
    try {
        const finishedPromise = new Promise((resolve, reject) => {
            chart.on('finished', resolve);
            chart.on('error', reject);
        });
        chart.setOption(drawing.renderOption);
        await finishedPromise;
        if (drawing.rotation) return _renderRotatedCanvasToDataURL(canvas, logicalW, logicalH, drawing.rotation, pixelRatio);
        return canvas.toDataURL('image/png');
    } finally {
        chart.dispose();
    }
}

function _renderShapeDrawingToDataURL(drawing, pixelRatio = 2) {
    const logicalW = Math.max(1, Number(drawing.width) || 1);
    const logicalH = Math.max(1, Number(drawing.height) || 1);
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(logicalW * pixelRatio));
    canvas.height = Math.max(1, Math.round(logicalH * pixelRatio));
    const ctx = canvas.getContext('2d');
    ctx.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
    if (drawing.rotation) {
        ctx.translate(logicalW / 2, logicalH / 2);
        ctx.rotate(drawing.rotation * Math.PI / 180);
        drawShapeToCanvas(ctx, drawing, -logicalW / 2, -logicalH / 2, logicalW, logicalH, pixelRatio);
    } else {
        drawShapeToCanvas(ctx, drawing, 0, 0, logicalW, logicalH, pixelRatio);
    }
    return canvas.toDataURL('image/png');
}

function _downloadDataURL(dataURL, filename) {
    const a = document.createElement('a');
    a.href = dataURL;
    a.download = filename;
    a.click();
    a.remove();
}

function _resolveDrawingDownloadName(SN, drawing, fallbackExt = 'png') {
    const workbookName = SN.workbookName || 'SheetNext';
    const type = drawing?.type || 'drawing';
    return `${workbookName}_${type}_${drawing?.id || 'item'}.${fallbackExt}`;
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

/** Download active drawing as image. */
export async function downloadActiveDrawingImage() {
    const SN = this.SN;
    const drawing = SN.Canvas?.activeDrawing;
    if (!drawing) return;

    const isChart = drawing.type === 'chart';
    const isShape = drawing.type === 'shape';
    const isImage = drawing.type === 'image';
    if (!isChart && !isShape && !isImage) {
        SN.Utils.toast(SN.t('action.insert.toast.downloadDrawingImageUnsupported'));
        return;
    }

    try {
        const imageBase64 = isImage
            ? drawing.imageBase64
            : isChart
            ? await _renderChartDrawingToDataURL(drawing)
            : _renderShapeDrawingToDataURL(drawing);
        const ext = isImage ? (drawing.imageSuffix || 'png') : 'png';
        _downloadDataURL(imageBase64, _resolveDrawingDownloadName(SN, drawing, ext));
        SN.Utils.toast(SN.t('action.insert.toast.downloadDrawingImageSuccess'));
    } catch (error) {
        console.error('downloadActiveDrawingImage failed:', error);
        SN.Utils.toast(error?.message || SN.t('action.insert.toast.downloadDrawingImageFailed'));
    }
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
                const selectedItem = getSelectedChartItem(bodyEl);
                if (selectedItem) {
                    this.insertChart(selectedItem);
                }
            }
        });
    } catch (e) {
        // 用户取消
    }
}

export function insertChart(chartType, options = {}) {
    const sheet = this.SN.activeSheet;
    const areas = sheet?.activeAreas;

    if (!areas || areas.length === 0) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.selectRange'));
    }

    const area = areas[areas.length - 1];
    const selectedType = chartType || this.SN.chartSelection?.item;
    const meta = resolveChartMeta(this.SN, selectedType);
    if (!meta) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.noType'));
    }
    if (!meta.supported) {
        return this.SN.Utils.toast(this.SN.t('chart.toast.unsupportedType', { label: meta.label }));
    }

    const result = createChartOptionFromRange(this.SN, sheet, area, selectedType, options);
    if (!result?.option) return this.SN.Utils.toast(this.SN.t('chart.toast.rangeEmpty'));

    const drawing = sheet.Drawing.addChart(result.option, options.drawing);
    if (!drawing) return false;

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
        cellRef: { r: targetCell.r, c: targetCell.c },
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
                    cellRef: { r: cellObj.r, c: cellObj.c },
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
                        outline: none; border-color: var(--sn-primary); box-shadow: 0 0 0 2px var(--sn-primary-ring);
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

function normalizeSheetName(name) {
    const text = String(name ?? '');
    return /[\s']/.test(text) ? `'${text.replaceAll("'", "''")}'` : text;
}

function buildRangeRef(sheet, sr, sc, er, ec) {
    const sheetName = normalizeSheetName(sheet.name);
    const start = sheet.Utils.cellNumToStr({ r: sr, c: sc });
    const end = sheet.Utils.cellNumToStr({ r: er, c: ec });
    return `${sheetName}!${start}:${end}`;
}
