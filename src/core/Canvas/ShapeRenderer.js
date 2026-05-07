// 形状绘制辅助函数
// 在指定 context 上绘制形状（供 SheetRenderer 调用）

import { shapePaths } from './shapePathsCanvas.js';
import { CONNECTOR_TYPES } from '../Drawing/constants.js';

export function drawShapeToCanvas(ctx, drawing, x, y, w, h, pixelRatio = 1) {
    const shapeType = drawing.shapeType || 'rect';
    const style = drawing.shapeStyle || {};
    const isConnector = drawing.isConnector || CONNECTOR_TYPES.includes(shapeType);
    const ratio = Number.isFinite(pixelRatio) && pixelRatio > 0 ? pixelRatio : 1;
    const snap = (value) => Math.round(value * ratio) / ratio;

    ctx.save();

    // 填充部分（连接线跳过填充）
    if (!isConnector && style.fill && style.fill !== 'none') {
        ctx.beginPath();
        const drawPath = shapePaths[shapeType];
        if (drawPath) {
            drawPath(ctx, x, y, w, h);
        } else {
            shapePaths.rect(ctx, x, y, w, h);
        }
        ctx.fillStyle = style.fill;
        ctx.fill();
    }

    // 边框部分（需要0.5偏移以锐化线条）
    const baseStrokeWidth = Math.max(0, Number(style.strokeWidth) || 0);
    const strokeWidth = !isConnector && drawing.isTextBox
        ? Math.max(0.5, baseStrokeWidth * 0.85)
        : baseStrokeWidth;
    if (style.stroke && strokeWidth > 0) {
        ctx.save();

        // 奇数像素宽度的线条需要0.5偏移才能锐化
        const deviceStrokeWidth = Math.max(1, Math.round(strokeWidth * ratio));
        const needsHalfPixelOffset = deviceStrokeWidth % 2 === 1;

        if (needsHalfPixelOffset) {
            const halfOffset = 0.5 / ratio;
            ctx.translate(halfOffset, halfOffset);
        }

        ctx.beginPath();
        const drawPath = shapePaths[shapeType];
        if (drawPath) {
            drawPath(ctx, x, y, w, h);
        } else {
            shapePaths.rect(ctx, x, y, w, h);
        }

        ctx.strokeStyle = style.stroke;
        ctx.lineWidth = strokeWidth;

        // 连接线用圆头，形状用方头
        if (isConnector) {
            ctx.lineCap = 'round';
            ctx.lineJoin = 'round';
        } else {
            ctx.lineCap = 'square';
            ctx.lineJoin = 'miter';
        }
        ctx.stroke();

        // 绘制箭头（仅连接线）
        if (isConnector) {
            drawConnectorArrows(ctx, shapeType, x, y, w, h, style);
        }

        ctx.restore();
    }

    ctx.restore();

    // 绘制文本（如果有，连接线通常没有文本）
    if (drawing.shapeText && !isConnector) {
        ctx.save();
        ctx.fillStyle = style.textColor || '#000000';
        ctx.font = `${style.fontSize || 11}px ${style.fontFamily || '宋体'}`;
        ctx.textBaseline = 'middle';
        const textDirection = style.textDirection === 'vertical' ? 'vertical' : 'horizontal';
        const isTextBox = !!drawing.isTextBox;

        // 文本对齐映射：Excel -> Canvas
        const alignMap = { 'l': 'left', 'r': 'right', 'ctr': 'center' };
        const defaultAlign = textDirection === 'vertical' ? 'right' : 'left';
        const canvasAlign = alignMap[style.textAlign] || (isTextBox ? defaultAlign : 'center');
        const lines = drawing.shapeText.split('\n');
        const lineHeight = (style.fontSize || 11) * 1.2;

        if (textDirection === 'vertical') {
            drawVerticalShapeText(ctx, lines, x, y, w, h, canvasAlign, lineHeight, isTextBox, ratio);
        } else {
            ctx.textAlign = canvasAlign;

            // 根据对齐方式计算 x 坐标
            const textX = canvasAlign === 'left' ? x + 5 :
                        canvasAlign === 'right' ? x + w - 5 :
                        x + w / 2;

            const startY = isTextBox
                ? y + 5 + lineHeight / 2
                : y + h / 2 - (lines.length * lineHeight) / 2 + lineHeight / 2;

            lines.forEach((line, i) => {
                ctx.fillText(line, snap(textX), snap(startY + i * lineHeight));
            });
        }

        ctx.restore();
    }
}

function drawVerticalShapeText(ctx, lines, x, y, w, h, align, lineHeight, isTextBox = false, pixelRatio = 1) {
    const ratio = Number.isFinite(pixelRatio) && pixelRatio > 0 ? pixelRatio : 1;
    const snap = (value) => Math.round(value * ratio) / ratio;
    const innerPadding = isTextBox ? 5 : 0;
    const maxCharsPerColumn = isTextBox
        ? Math.max(1, Math.floor(Math.max(lineHeight, h - innerPadding * 2) / lineHeight))
        : Number.MAX_SAFE_INTEGER;
    const columns = [];
    lines.forEach((line) => {
        const chars = Array.from(line || '');
        if (chars.length === 0) {
            columns.push([]);
            return;
        }
        for (let i = 0; i < chars.length; i += maxCharsPerColumn) {
            columns.push(chars.slice(i, i + maxCharsPerColumn));
        }
    });
    if (columns.length === 0) columns.push([]);
    const columnCount = Math.max(columns.length, 1);
    const maxChars = Math.max(...columns.map(chars => chars.length), 1);
    const totalWidth = columnCount * lineHeight;
    const totalHeight = maxChars * lineHeight;

    const startX = align === 'left' ? x + innerPadding :
        align === 'right' ? x + w - innerPadding - totalWidth :
        x + (w - totalWidth) / 2;
    const startY = isTextBox ? y + innerPadding : y + (h - totalHeight) / 2;

    // Excel 竖排文本框默认按列从右向左排列
    if (isTextBox) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(x + innerPadding, y + innerPadding, Math.max(0, w - innerPadding * 2), Math.max(0, h - innerPadding * 2));
        ctx.clip();
    }
    ctx.textAlign = 'center';
    columns.forEach((chars, lineIndex) => {
        const col = columnCount - 1 - lineIndex;
        const textX = startX + col * lineHeight + lineHeight / 2;
        chars.forEach((char, charIndex) => {
            const textY = startY + charIndex * lineHeight + lineHeight / 2;
            ctx.fillText(char, snap(textX), snap(textY));
        });
    });
    if (isTextBox) ctx.restore();
}

// 绘制连接线箭头
function drawConnectorArrows(ctx, shapeType, x, y, w, h, style) {
    const arrowSize = Math.max(8, style.strokeWidth * 3);

    // 获取线条的起点和终点方向
    const { startPoint, endPoint, startAngle, endAngle } = getConnectorEndpoints(shapeType, x, y, w, h);

    // 绘制起始箭头
    if (style.startArrow && style.startArrow !== 'none') {
        drawArrowHead(ctx, startPoint.x, startPoint.y, startAngle + Math.PI, arrowSize, style);
    }

    // 绘制终点箭头
    if (style.endArrow && style.endArrow !== 'none') {
        drawArrowHead(ctx, endPoint.x, endPoint.y, endAngle, arrowSize, style);
    }
}

// 获取连接线端点和方向（仅4种标准连接线类型）
function getConnectorEndpoints(shapeType, x, y, w, h) {
    let startPoint, endPoint, startAngle, endAngle;

    switch (shapeType) {
        case 'line':
        case 'straightConnector1':
            // 对角线：从左上到右下
            startPoint = { x: x, y: y };
            endPoint = { x: x + w, y: y + h };
            startAngle = Math.atan2(-h, -w);
            endAngle = Math.atan2(h, w);
            break;

        case 'bentConnector3':
            // 肘形连接符：起点向右，终点向右
            startPoint = { x: x, y: y };
            endPoint = { x: x + w, y: y + h };
            startAngle = 0; // 起点向右
            endAngle = 0; // 终点向右
            break;

        case 'curvedConnector3':
            // 曲线连接符：起点向右，终点向右
            startPoint = { x: x, y: y };
            endPoint = { x: x + w, y: y + h };
            startAngle = 0; // 起点向右
            endAngle = 0; // 终点向右
            break;

        default:
            // 默认对角线
            startPoint = { x: x, y: y };
            endPoint = { x: x + w, y: y + h };
            startAngle = Math.atan2(-h, -w);
            endAngle = Math.atan2(h, w);
    }

    return { startPoint, endPoint, startAngle, endAngle };
}

// 绘制箭头头部
function drawArrowHead(ctx, x, y, angle, size, style) {
    ctx.save();
    ctx.translate(x, y);
    ctx.rotate(angle);

    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(-size, -size / 2);
    ctx.lineTo(-size, size / 2);
    ctx.closePath();

    ctx.fillStyle = style.stroke || '#000000';
    ctx.fill();

    ctx.restore();
}
