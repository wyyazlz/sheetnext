// 形状绘制辅助函数
// 在指定 context 上绘制形状（供 SheetRenderer 调用）

import { getShapeGeometry } from './shapePathsCanvas.js';
import { drawConnector, drawArrowHead } from './connectorPathsCanvas.js';
import { CONNECTOR_TYPES } from '../Drawing/constants.js';

export function drawShapeToCanvas(ctx, drawing, x, y, w, h, pixelRatio = 1) {
    const shapeType = drawing.shapeType || 'rect';
    const style = drawing.shapeStyle || {};
    const isConnector = drawing.isConnector || CONNECTOR_TYPES.includes(shapeType);
    const ratio = Number.isFinite(pixelRatio) && pixelRatio > 0 ? pixelRatio : 1;
    const snap = (value) => Math.round(value * ratio) / ratio;
    const geometry = isConnector ? null : getShapeGeometry(drawing, w, h);

    ctx.save();
    const baseStrokeWidth = Math.max(0, Number(style.strokeWidth) || 0);
    const strokeWidth = !isConnector && drawing.isTextBox && baseStrokeWidth > 0
        ? Math.max(0.5, baseStrokeWidth * 0.85)
        : baseStrokeWidth;
    if (isConnector) {
        if (style.stroke && style.stroke !== 'none' && strokeWidth > 0) {
            const angles = drawConnector(ctx, shapeType, x, y, x + w, y + h, style);
            const size = strokeWidth * 8;
            drawArrowHead(ctx, x, y, angles.startAngle + Math.PI, size, style.startArrow, style.stroke);
            drawArrowHead(ctx, x + w, y + h, angles.endAngle, size, style.endArrow, style.stroke);
        }
    } else {
        ctx.translate(x, y);
        ctx.lineWidth = strokeWidth;
        ctx.lineCap = 'butt';
        ctx.lineJoin = 'miter';
        ctx.miterLimit = 10;
        ctx.setLineDash([]);
        const hasStroke = style.stroke && style.stroke !== 'none' && strokeWidth > 0;
        const deviceStrokeWidth = Math.max(1, Math.round(strokeWidth * ratio));
        const offset = hasStroke && deviceStrokeWidth % 2 === 1 ? 0.5 / ratio : 0;
        ctx.translate(offset, offset);
        for (const layer of geometry.paths) {
            if (layer.fill !== 'none' && style.fill && style.fill !== 'none') {
                ctx.fillStyle = style.fill;
                if (layer.fill !== 'norm') ctx.fillStyle = shadeShapeColor(ctx.fillStyle, layer.fill);
                ctx.fill(layer.path);
            }
            if (layer.stroke && hasStroke) {
                ctx.strokeStyle = style.stroke;
                ctx.stroke(layer.path);
            }
        }
    }

    ctx.restore();

    // 绘制文本（如果有，连接线通常没有文本）
    if (drawing.shapeText && !isConnector) {
        ctx.save();
        if (!drawing.isTextBox) {
            x += geometry.textRect.x;
            y += geometry.textRect.y;
            w = geometry.textRect.w;
            h = geometry.textRect.h;
        }
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

function shadeShapeColor(color, mode) {
    const hex = /^#([\da-f]{6})([\da-f]{2})?$/i.exec(color);
    const rgba = /^rgba?\(([^)]+)\)$/.exec(color);
    if (!hex && !rgba) return color;
    const channels = hex
        ? [0, 2, 4].map(offset => parseInt(hex[1].slice(offset, offset + 2), 16))
        : rgba[1].split(',').map(Number);
    const alpha = hex ? (hex[2] ? parseInt(hex[2], 16) / 255 : 1) : (channels[3] ?? 1);
    const amount = mode.endsWith('Less') ? 0.2 : 0.4;
    const light = mode.startsWith('lighten');
    const rgb = channels.slice(0, 3).map(value => Math.round(light ? value + (255 - value) * amount : value * (1 - amount)));
    return `rgba(${rgb.join(',')},${alpha})`;
}
