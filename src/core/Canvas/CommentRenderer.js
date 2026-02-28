/**
 * 批注渲染模块
 * 负责渲染批注指示器（红色三角）和批注浮动框
 * 高性能设计：仅渲染视图内的批注
 */

import { getIndicatorPosition } from '../Comment/helpers.js';

/**
 * 渲染批注指示器
 * @param {Sheet} sheet - 工作表实例
 */
export function rCommentIndicators(sheet) {
    const comments = sheet.Comment.getAll();
    if (comments.length === 0) return;

    const ctx = this.bc; // 使用缓冲画布上下文

    // 只渲染视图内的批注
    const visibleComments = comments.filter(comment => {
        const { r, c } = comment.cellPos;
        const cellInfo = sheet.getCellInViewInfo(r, c);
        return cellInfo && cellInfo.inView; // 检查inView属性
    });

    ctx.save();
    ctx.fillStyle = '#FF0000'; // 红色三角

    visibleComments.forEach(comment => {
        const { r, c } = comment.cellPos;
        const cellInfo = sheet.getCellInViewInfo(r, c);

        if (!cellInfo || !cellInfo.inView) return;

        const { x, y, size } = getIndicatorPosition(cellInfo);

        // 绘制红色三角（右上角）
        ctx.beginPath();
        ctx.moveTo(x + size, y); // 右上角点
        ctx.lineTo(x + size, y + size); // 右下点
        ctx.lineTo(x, y); // 左上点
        ctx.closePath();
        ctx.fill();
    });

    ctx.restore();
}

/**
 * 渲染批注浮动框
 * @param {Sheet} sheet - 工作表实例
 */
export function rCommentBoxes(sheet) {
    const ctx = this.bc;

    // 获取需要显示的批注（visible=true 或 hoveredComment）
    const commentsToShow = [];

    // 添加所有visible的批注
    sheet.Comment.getAll().forEach(comment => {
        if (comment.visible) {
            commentsToShow.push(comment);
        }
    });

    // 添加hoveredComment（如果存在且不重复）
    if (this.hoveredComment && !this.hoveredComment.visible) {
        commentsToShow.push(this.hoveredComment);
    }

    if (commentsToShow.length === 0) return;

    ctx.save();

    commentsToShow.forEach(comment => {
        const { r, c } = comment.cellPos;
        const cellInfo = sheet.getCellInViewInfo(r, c);

        if (!cellInfo || !cellInfo.inView) return; // 检查inView属性

        // 计算批注框位置（相对于单元格）
        const pxPerPt = 1.33;
        const x = cellInfo.x + comment.marginLeft * pxPerPt;
        const y = cellInfo.y + comment.marginTop * pxPerPt;
        const width = comment.width * pxPerPt;
        const height = comment.height * pxPerPt;

        // 1. 绘制阴影
        ctx.shadowColor = 'rgba(0, 0, 0, 0.2)';
        ctx.shadowBlur = 4;
        ctx.shadowOffsetX = 2;
        ctx.shadowOffsetY = 2;

        // 2. 绘制批注框背景
        ctx.fillStyle = '#FFFFE1';
        ctx.fillRect(x, y, width, height);

        // 3. 绘制边框
        ctx.shadowColor = 'transparent'; // 清除阴影
        ctx.strokeStyle = '#000000';
        ctx.lineWidth = 1;
        ctx.strokeRect(x, y, width, height);

        // 4. 绘制连接线和小三角（指向单元格右上角）
        const indicatorPos = getIndicatorPosition(cellInfo);
        const arrowEndX = indicatorPos.x + indicatorPos.size;
        const arrowEndY = indicatorPos.y;

        // 绘制连接线
        ctx.strokeStyle = '#000000';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(x, y);
        ctx.lineTo(arrowEndX, arrowEndY);
        ctx.stroke();

        // 绘制单元格端小三角（指向单元格）
        const triSize = 5;
        const dx = x - arrowEndX;
        const dy = y - arrowEndY;
        const len = Math.sqrt(dx * dx + dy * dy);
        const ux = dx / len;
        const uy = dy / len;
        // 垂直方向
        const px = -uy;
        const py = ux;

        ctx.fillStyle = '#000000';
        ctx.beginPath();
        ctx.moveTo(arrowEndX, arrowEndY); // 箭头尖端
        ctx.lineTo(arrowEndX + ux * triSize * 2 + px * triSize, arrowEndY + uy * triSize * 2 + py * triSize);
        ctx.lineTo(arrowEndX + ux * triSize * 2 - px * triSize, arrowEndY + uy * triSize * 2 - py * triSize);
        ctx.closePath();
        ctx.fill();

        // 5. 绘制文本内容
        if (comment.text) {
            ctx.fillStyle = '#000000';

            // 多行文本渲染
            const paddingX = 10;
            const paddingY = 0;
            const lineHeight = 16;
            const maxWidth = width - paddingX * 2;

            let startY = y + paddingY + 12;

            // 绘制作者名（如果有）
            if (comment.author) {
                ctx.font = 'bold 12px Arial';
                ctx.fillText(comment.author + ':', x + paddingX, startY);
                startY += lineHeight;
            }

            // 绘制批注内容
            ctx.font = '12px Arial';
            const lines = wrapText(ctx, comment.text, maxWidth);

            lines.forEach((line, index) => {
                const lineY = startY + index * lineHeight;
                if (lineY + lineHeight <= y + height - paddingY) { // 不超出边界
                    ctx.fillText(line, x + paddingX, lineY);
                }
            });
        }
    });

    ctx.restore();
}

/**
 * 文本换行处理
 * @private
 */
function wrapText(ctx, text, maxWidth) {
    const lines = [];
    const paragraphs = text.split('\n');

    paragraphs.forEach(paragraph => {
        if (!paragraph) {
            lines.push('');
            return;
        }

        let line = '';
        const words = paragraph.split('');

        for (let i = 0; i < words.length; i++) {
            const testLine = line + words[i];
            const metrics = ctx.measureText(testLine);

            if (metrics.width > maxWidth && line) {
                lines.push(line);
                line = words[i];
            } else {
                line = testLine;
            }
        }

        if (line) {
            lines.push(line);
        }
    });

    return lines;
}
