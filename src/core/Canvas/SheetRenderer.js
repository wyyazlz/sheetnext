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
import { drawShapeToCanvas } from './ShapeRenderer.js';
import { drawConnector, drawArrowHead } from './connectorPathsCanvas.js';

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

/**
 * Drawing Rendering Module
 * Responsible for handling the rendering logic of drawings, including pictures, diagrams, etc.
 */

// 通过定位信息获取对应坐标和偏移量
/** @param {number} x @param {number} y @param {number} w @param {number} h @returns {Object} */
export function getOffset(x, y, w, h) {
    const rect = this.handleLayer.getBoundingClientRect();
    const zoom = this.activeSheet?.zoom ?? 1;
    // 先获取单元格范围
    const obj = { s: {}, e: {} }
    const { r: sr, c: sc } = this.getCellIndex({ clientX: rect.left + x * zoom, clientY: rect.top + y * zoom });
    const { r: er, c: ec } = this.getCellIndex({ clientX: rect.left + (x + w) * zoom, clientY: rect.top + (y + h) * zoom });
    obj.s.r = sr
    obj.s.c = sc
    obj.e.r = er
    obj.e.c = ec
    // 再获取偏移
    const sheet = this.activeSheet
    const startInfo = sheet.getCellInViewInfo(sr, sc, false);
    const endInfo = sheet.getCellInViewInfo(er, ec, false);
    obj.s.offsetX = x - startInfo.x;
    obj.s.offsetY = y - startInfo.y;
    obj.e.offsetX = x + w - endInfo.x;
    obj.e.offsetY = y + h - endInfo.y;
    return obj;
}

// 渲染图纸(drawings)
/** @param {number} nowRCount @param {Object} [options={}] @returns {Promise<void>} */
export async function rDrawing(nowRCount, options = {}) {

    const isScreenshot = !!options.targetCanvas;
    if (!isScreenshot && this.slicersCon) {
        this.slicersCon.querySelectorAll('.sn-slicer').forEach(el => {
            el.style.display = 'none';
        });
    }

    // 截图模式不操作handleLayer
    if (!isScreenshot) this.HDctx.clearRect(0, 0, this.showLayer.width, this.showLayer.height);

    const sheet = this.activeSheet
    if (sheet.Drawing.getAll().length == 0) {
        this._syncTextBoxEditorPosition?.();
        return;
    }

    // 截图模式不清除buffer（在数据层上叠加图纸）
    if (!isScreenshot) this.clearBuffer(); // 清除缓冲画布

    // 遍历所有 drawing
    for (let drawing of sheet.Drawing.getAll()) {

        // 等待图片初始化完成（获取尺寸+缓存图片）
        if (drawing._initPromise) {
            await drawing._initPromise;
            if (nowRCount !== this.rCount) return; // 初始化期间可能触发新渲染
        }

        const info = sheet._getAreaInviewInfo2(drawing.area, true); // 是否在当前视图内
        drawing.inView = info.inView
        if (!info.inView) {
            if (drawing.type === 'slicer' && drawing.Slicer) {
                drawing.Slicer.render(this, 0, 0, 0, 0, false);
            }
            continue;
        }

        // 进来时应该直接计算长高,否则拉列宽时会影响
        // 计算图纸在单元格的偏移信息,实时变化；隐藏行列在屏幕上占 0px，锚点行列被隐藏时偏移量按 0 计
        const sCol = sheet.getCol(drawing.area.s.c);
        const sRow = sheet.getRow(drawing.area.s.r);
        const eCol = sheet.getCol(drawing.area.e.c);
        const eRow = sheet.getRow(drawing.area.e.r);
        const sOffX = sCol.hidden ? 0 : drawing.area.s.offsetX;
        const sOffY = sRow.hidden ? 0 : drawing.area.s.offsetY;
        let x, y
        let w = info.w - sOffX - (eCol.hidden ? 0 : eCol.width - drawing.area.e.offsetX);
        let h = info.h - sOffY - (eRow.hidden ? 0 : eRow.height - drawing.area.e.offsetY);
        if (drawing.anchorType !== 'twoCell') {
            w = drawing.width;
            h = drawing.height;
        }

        if (drawing.moving) { // 都move了肯定是有position信息的
            x = drawing.position.x
            y = drawing.position.y
        } else {
            x = info.x + sOffX
            y = info.y + sOffY
        }

        if (drawing.type === 'slicer') {
            if (drawing.Slicer) {
                drawing.Slicer.render(this, x, y, w, h, true);
            }
        } else if (drawing.type === 'image') {
            try {
                // 如果 drawCache 为空，重新加载图片
                if (!drawing.drawCache) {
                    const img = new Image();
                    await new Promise((resolve, reject) => {
                        img.onload = () => {
                            drawing.drawCache = img;
                            resolve();
                        };
                        img.onerror = reject;
                        img.src = drawing.imageBase64;
                    });
                    if (nowRCount !== this.rCount) return;
                }
                // 绘制前再检查
                if (nowRCount !== this.rCount) return;
                if (drawing._rotation) {
                    this.bc.save();
                    this.bc.translate(x + w / 2, y + h / 2);
                    this.bc.rotate(drawing._rotation * Math.PI / 180);
                    this.bc.drawImage(drawing.drawCache, -w / 2, -h / 2, w, h);
                    this.bc.restore();
                } else {
                    this.bc.drawImage(drawing.drawCache, x, y, w, h);
                }
            } catch (error) {
                console.error('图片加载失败:', error);
            }
        }

        if (drawing.type === 'chart') {
            const pixelRatio = this.pixelRatio;
            const rect = this.snapRect(x, y, w, h);
            const logicalW = rect.w;
            const logicalH = rect.h;
            const needsRender = !drawing.drawCache ||
                drawing._cachedWidth !== logicalW ||
                drawing._cachedHeight !== logicalH ||
                !Number.isFinite(drawing._cachedPixelRatio) ||
                Math.abs(drawing._cachedPixelRatio - pixelRatio) > 1e-3;

            if (needsRender) {
                // 复用 Canvas
                const localCount = this.rCount;
                const canvas = document.createElement('canvas');
                const physicalW = Math.max(1, Math.round(logicalW * pixelRatio));
                const physicalH = Math.max(1, Math.round(logicalH * pixelRatio));
                canvas.width = physicalW;
                canvas.height = physicalH;
                const chart = echarts.init(canvas, null, {
                    renderer: 'canvas',
                    width: logicalW,
                    height: logicalH,
                    devicePixelRatio: pixelRatio
                });
                drawing.drawCache = canvas;
                drawing._cachedWidth = logicalW;
                drawing._cachedHeight = logicalH;
                drawing._cachedPixelRatio = pixelRatio;
                // 先监听再渲染
                const finishedPromise = new Promise((resolve, reject) => {
                    chart.on('finished', () => {
                        resolve();
                    });
                    chart.on('error', err => {
                        reject(err);
                    });
                });
                chart.setOption(drawing.renderOption);
                await finishedPromise;

                if (localCount !== this.rCount) return;
            }

            // 绘制前再检查
            if (nowRCount !== this.rCount) return;
            const alignOffset = this.halfPixel;
            const prevSmoothing = this.bc.imageSmoothingEnabled;
            this.bc.imageSmoothingEnabled = false;
            this.bc.drawImage(drawing.drawCache, rect.x + alignOffset, rect.y + alignOffset, rect.w, rect.h);
            this.bc.imageSmoothingEnabled = prevSmoothing;
            this.bc.strokeStyle = '#ccc';
            this.bc.lineWidth = this.px(1);
            this.bc.strokeRect(rect.x + alignOffset, rect.y + alignOffset, rect.w, rect.h);
        }

        if (drawing.type === 'shape') {

            // 连接线渲染
            if (drawing.isConnector) {
                // 连接线：从起点到终点绘制
                const x1 = x;
                const y1 = y;
                const x2 = x + w;
                const y2 = y + h;

                const style = drawing.shapeStyle;
                const connectorType = drawing.shapeType || 'straightConnector1';

                // 绘制连接线路径，返回起点和终点的角度
                const angles = drawConnector(this.bc, connectorType, x1, y1, x2, y2, style);

                // 绘制箭头
                const arrowSize = (style.strokeWidth || 1) * 8;

                // 绘制起点箭头（方向相反）
                if (style.startArrow && style.startArrow !== 'none') {
                    drawArrowHead(this.bc, x1, y1, angles.startAngle + Math.PI, arrowSize, style.startArrow, style.stroke);
                }

                // 绘制终点箭头
                if (style.endArrow && style.endArrow !== 'none') {
                    drawArrowHead(this.bc, x2, y2, angles.endAngle, arrowSize, style.endArrow, style.stroke);
                }
            } else {
                // 普通形状渲染
                const pixelRatio = this.pixelRatio;
                const logicalW = Math.max(1, w);
                const logicalH = Math.max(1, h);

                if (drawing.isTextBox) {
                    const directRect = this.snapRect(x, y, logicalW, logicalH);
                    if (drawing._rotation) {
                        this.bc.save();
                        this.bc.translate(directRect.x + directRect.w / 2, directRect.y + directRect.h / 2);
                        this.bc.rotate(drawing._rotation * Math.PI / 180);
                        drawShapeToCanvas(this.bc, drawing, -directRect.w / 2, -directRect.h / 2, directRect.w, directRect.h, pixelRatio);
                        this.bc.restore();
                    } else {
                        drawShapeToCanvas(this.bc, drawing, directRect.x, directRect.y, directRect.w, directRect.h, pixelRatio);
                    }
                } else {
                    const strokeWidth = Math.max(0, drawing.shapeStyle?.strokeWidth || 0);
                    const padding = Math.max(1, Math.ceil(strokeWidth * 2));
                    const canvasW = logicalW + padding * 2;
                    const canvasH = logicalH + padding * 2;

                    const needsRender = !drawing.drawCache ||
                        drawing._cachedWidth !== logicalW ||
                        drawing._cachedHeight !== logicalH ||
                        drawing._cachedStrokeWidth !== strokeWidth ||
                        !Number.isFinite(drawing._cachedPixelRatio) ||
                        Math.abs(drawing._cachedPixelRatio - pixelRatio) > 1e-3;

                    if (needsRender) {
                        const canvas = document.createElement('canvas');
                        canvas.width = Math.max(1, Math.round(canvasW * pixelRatio));
                        canvas.height = Math.max(1, Math.round(canvasH * pixelRatio));

                        const ctx = canvas.getContext('2d');
                        ctx.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
                        drawShapeToCanvas(ctx, drawing, padding, padding, logicalW, logicalH, pixelRatio);

                        drawing.drawCache = canvas;
                        drawing._cachedWidth = logicalW;
                        drawing._cachedHeight = logicalH;
                        drawing._cachedStrokeWidth = strokeWidth;
                        drawing._cachedPixelRatio = pixelRatio;
                    }

                    // render cancellation check
                    if (nowRCount !== this.rCount) return;

                    const drawX = x - padding;
                    const drawY = y - padding;
                    const drawW = logicalW + padding * 2;
                    const drawH = logicalH + padding * 2;
                    const snappedRect = this.snapRect(drawX, drawY, drawW, drawH);

                    if (drawing._rotation) {
                        this.bc.save();
                        this.bc.translate(snappedRect.x + snappedRect.w / 2, snappedRect.y + snappedRect.h / 2);
                        this.bc.rotate(drawing._rotation * Math.PI / 180);
                        this.bc.drawImage(drawing.drawCache, -snappedRect.w / 2, -snappedRect.h / 2, snappedRect.w, snappedRect.h);
                        this.bc.restore();
                    } else {
                        this.bc.drawImage(drawing.drawCache, snappedRect.x, snappedRect.y, snappedRect.w, snappedRect.h);
                    }
                }
            }

        }

        if (drawing.active) {
            // 正方形尺寸,可根据需要调整
            const handleSize = 8;
            const halfSize = handleSize / 2;

            // 计算中心点
            const cx = x + w / 2;
            const cy = y + h / 2;

            // 如果有旋转，将选中框也旋转绘制
            if (drawing._rotation) {
                this.bc.save();
                this.bc.translate(cx, cy);
                this.bc.rotate(drawing._rotation * Math.PI / 180);
                this.bc.translate(-cx, -cy);
            }

            // 8 个点：左上、上中、右上、左中、右中、左下、下中、右下
            const points = [
                { x: x, y: y },         // 左上
                { x: cx, y: y },         // 上中
                { x: x + w, y: y },         // 右上
                { x: x, y: cy },        // 左中
                { x: x + w, y: cy },        // 右中
                { x: x, y: y + h },     // 左下
                { x: cx, y: y + h },     // 下中
                { x: x + w, y: y + h }      // 右下
            ];

            // 样式：白底,边框 #333
            this.bc.fillStyle = 'white';
            this.bc.strokeStyle = '#333';
            this.bc.lineWidth = 1;
            this.bc.strokeRect(x, y, w, h);
            // 依次绘制每个拖拽点为正方形
            points.forEach(pt => {
                // 以 pt.x, pt.y 作为正方形中心
                const px = pt.x - halfSize;
                const py = pt.y - halfSize;
                this.bc.fillRect(px, py, handleSize, handleSize);
                this.bc.strokeRect(px, py, handleSize, handleSize);
            });

            // 在 active 绘制完八个拖拽点之后,紧接着加：
            if (drawing.type === 'image' || (drawing.type === 'shape' && !drawing.isConnector)) {
                // 旋转按钮样式
                const handleRadius = 5;           // 圆半径,可根据需要调整
                const offsetY = 30;               // 圆心到边框顶部的距离
                const handleX = cx;               // 顶部水平居中
                const handleY = y - offsetY;      // 在矩形上方

                // 连接线：从顶部中点连到旋转按钮
                this.bc.strokeStyle = '#333';
                this.bc.beginPath();
                this.bc.moveTo(cx, y - halfSize);                 // 矩形顶部中点
                this.bc.lineTo(handleX, handleY + handleRadius);
                this.bc.stroke();

                // 绘制圆形旋转按钮
                this.bc.beginPath();
                this.bc.arc(handleX, handleY, handleRadius, 0, Math.PI * 2);
                this.bc.fillStyle = 'white';
                this.bc.fill();
                this.bc.stroke();
            }

            if (drawing._rotation) {
                this.bc.restore();
            }
        }

        drawing.position = { x, y, w, h }

    }

    this._syncTextBoxEditorPosition?.();
    if (nowRCount !== this.rCount) return  // 重要,解决异步竞争渲染问题

    // 截图模式不操作handleLayer
    if (!isScreenshot) {
        // 将图纸从绘制层转到对应的显示层
        this.bc.clearRect(-0.5, -0.5, this.showLayer.width + 1, sheet.headHeight + 1)
        this.bc.clearRect(-0.5, -0.5, sheet.indexWidth + 1, this.showLayer.height + 1)
        this.HDctx.drawImage(this.buffer, 0, 0);
    }

}
