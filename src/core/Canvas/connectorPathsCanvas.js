// 连接线路径绘制函数集（高性能 Canvas 实现）
// 每个函数签名：(ctx, x1, y1, x2, y2) => void

/**
 * Connection line path mapping table
 * All cables are drawn from (x1, y1) to (x2, y2)
 * Returns {startAngle, endAngle} for arrow direction
 */
export const connectorPaths = {

    // ==================== 直线连接器 ====================

    // 基础直线（Excel 标准）
    line(ctx, x1, y1, x2, y2) {
        ctx.moveTo(x1, y1);
        ctx.lineTo(x2, y2);
        const angle = Math.atan2(y2 - y1, x2 - x1);
        return { startAngle: angle, endAngle: angle };
    },

    // 直线连接器（line 的别名）
    straightConnector1(ctx, x1, y1, x2, y2) {
        return connectorPaths.line(ctx, x1, y1, x2, y2);
    },

    // ==================== 折线连接器 (Bent Connectors) ====================

    // 两段折线：水平后垂直 或 垂直后水平
    bentConnector2(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        // 根据距离选择最佳转折方向
        if (Math.abs(dx) > Math.abs(dy)) {
            // 水平优先：水平 → 垂直
            const midX = x1 + dx / 2;
            ctx.moveTo(x1, y1);
            ctx.lineTo(midX, y1);
            ctx.lineTo(midX, y2);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        } else {
            // 垂直优先：垂直 → 水平
            const midY = y1 + dy / 2;
            ctx.moveTo(x1, y1);
            ctx.lineTo(x1, midY);
            ctx.lineTo(x2, midY);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        }
    },

    // 三段折线：水平 → 垂直 → 水平 或 垂直 → 水平 → 垂直
    bentConnector3(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        ctx.moveTo(x1, y1);

        if (Math.abs(dx) > Math.abs(dy)) {
            // 水平-垂直-水平
            const seg1 = dx * 0.5;
            ctx.lineTo(x1 + seg1, y1);
            ctx.lineTo(x1 + seg1, y2);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        } else {
            // 垂直-水平-垂直
            const seg1 = dy * 0.5;
            ctx.lineTo(x1, y1 + seg1);
            ctx.lineTo(x2, y1 + seg1);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        }
    },

    // 四段折线
    bentConnector4(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        ctx.moveTo(x1, y1);

        if (Math.abs(dx) > Math.abs(dy)) {
            // 水平-垂直-水平-垂直
            const seg1 = dx * 0.33;
            const seg2 = dy * 0.5;
            ctx.lineTo(x1 + seg1, y1);
            ctx.lineTo(x1 + seg1, y1 + seg2);
            ctx.lineTo(x2, y1 + seg2);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        } else {
            // 垂直-水平-垂直-水平
            const seg1 = dy * 0.33;
            const seg2 = dx * 0.5;
            ctx.lineTo(x1, y1 + seg1);
            ctx.lineTo(x1 + seg2, y1 + seg1);
            ctx.lineTo(x1 + seg2, y2);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        }
    },

    // 五段折线
    bentConnector5(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        ctx.moveTo(x1, y1);

        if (Math.abs(dx) > Math.abs(dy)) {
            // 水平-垂直-水平-垂直-水平
            const seg1 = dx * 0.25;
            const seg2 = dy * 0.5;
            const seg3 = dx * 0.5;
            ctx.lineTo(x1 + seg1, y1);
            ctx.lineTo(x1 + seg1, y1 + seg2);
            ctx.lineTo(x1 + seg3, y1 + seg2);
            ctx.lineTo(x1 + seg3, y2);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        } else {
            // 垂直-水平-垂直-水平-垂直
            const seg1 = dy * 0.25;
            const seg2 = dx * 0.5;
            const seg3 = dy * 0.5;
            ctx.lineTo(x1, y1 + seg1);
            ctx.lineTo(x1 + seg2, y1 + seg1);
            ctx.lineTo(x1 + seg2, y1 + seg3);
            ctx.lineTo(x2, y1 + seg3);
            ctx.lineTo(x2, y2);
            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        }
    },

    // ==================== 曲线连接器 (Curved Connectors) ====================

    // 简单曲线：单个贝塞尔曲线
    curvedConnector2(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        ctx.moveTo(x1, y1);

        if (Math.abs(dx) > Math.abs(dy)) {
            // 水平方向曲线
            const cpOffset = dx * 0.5;
            ctx.bezierCurveTo(
                x1 + cpOffset, y1,
                x2 - cpOffset, y2,
                x2, y2
            );
            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        } else {
            // 垂直方向曲线
            const cpOffset = dy * 0.5;
            ctx.bezierCurveTo(
                x1, y1 + cpOffset,
                x2, y2 - cpOffset,
                x2, y2
            );
            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        }
    },

    // 三段曲线：起点水平/垂直 → 转折 → 终点水平/垂直
    curvedConnector3(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        ctx.moveTo(x1, y1);

        if (Math.abs(dx) > Math.abs(dy)) {
            // 水平-转折-水平
            const midX = x1 + dx * 0.5;
            const smoothness = 0.25; // 平滑系数

            // 第一段：从起点水平出发
            const cp1x = x1 + dx * smoothness;
            const cp1y = y1;

            // 第二段：中间转折
            const cp2x = midX;
            const cp2y = y1;

            const cp3x = midX;
            const cp3y = y2;

            // 第三段：水平到达终点
            const cp4x = x2 - dx * smoothness;
            const cp4y = y2;

            ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, midX, y1 + dy * 0.5);
            ctx.bezierCurveTo(cp3x, cp3y, cp4x, cp4y, x2, y2);

            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        } else {
            // 垂直-转折-垂直
            const midY = y1 + dy * 0.5;
            const smoothness = 0.25;

            const cp1x = x1;
            const cp1y = y1 + dy * smoothness;

            const cp2x = x1;
            const cp2y = midY;

            const cp3x = x2;
            const cp3y = midY;

            const cp4x = x2;
            const cp4y = y2 - dy * smoothness;

            ctx.bezierCurveTo(cp1x, cp1y, cp2x, cp2y, x1 + dx * 0.5, midY);
            ctx.bezierCurveTo(cp3x, cp3y, cp4x, cp4y, x2, y2);

            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        }
    },

    // 四段曲线
    curvedConnector4(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        ctx.moveTo(x1, y1);

        if (Math.abs(dx) > Math.abs(dy)) {
            const seg1X = x1 + dx * 0.33;
            const seg2X = x1 + dx * 0.67;
            const midY = y1 + dy * 0.5;
            const smooth = dx * 0.15;

            // 起点 → 第一拐点
            ctx.bezierCurveTo(
                x1 + smooth, y1,
                seg1X - smooth, y1,
                seg1X, y1
            );

            // 第一拐点 → 中间
            ctx.bezierCurveTo(
                seg1X + smooth, y1,
                seg1X + smooth, midY,
                seg1X, midY
            );

            // 中间 → 第二拐点
            ctx.bezierCurveTo(
                seg1X, midY,
                seg2X, midY,
                seg2X, y2
            );

            // 第二拐点 → 终点
            ctx.bezierCurveTo(
                seg2X + smooth, y2,
                x2 - smooth, y2,
                x2, y2
            );

            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        } else {
            const seg1Y = y1 + dy * 0.33;
            const seg2Y = y1 + dy * 0.67;
            const midX = x1 + dx * 0.5;
            const smooth = dy * 0.15;

            ctx.bezierCurveTo(
                x1, y1 + smooth,
                x1, seg1Y - smooth,
                x1, seg1Y
            );

            ctx.bezierCurveTo(
                x1, seg1Y + smooth,
                midX, seg1Y + smooth,
                midX, seg1Y
            );

            ctx.bezierCurveTo(
                midX, seg1Y,
                midX, seg2Y,
                x2, seg2Y
            );

            ctx.bezierCurveTo(
                x2, seg2Y + smooth,
                x2, y2 - smooth,
                x2, y2
            );

            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        }
    },

    // 五段曲线
    curvedConnector5(ctx, x1, y1, x2, y2) {
        const dx = x2 - x1;
        const dy = y2 - y1;

        ctx.moveTo(x1, y1);

        if (Math.abs(dx) > Math.abs(dy)) {
            const points = [
                { x: x1, y: y1 },
                { x: x1 + dx * 0.25, y: y1 },
                { x: x1 + dx * 0.4, y: y1 + dy * 0.5 },
                { x: x1 + dx * 0.6, y: y1 + dy * 0.5 },
                { x: x1 + dx * 0.75, y: y2 },
                { x: x2, y: y2 }
            ];

            for (let i = 0; i < points.length - 1; i++) {
                const p1 = points[i];
                const p2 = points[i + 1];
                const cpDist = 0.3;

                ctx.bezierCurveTo(
                    p1.x + (p2.x - p1.x) * cpDist, p1.y,
                    p2.x - (p2.x - p1.x) * cpDist, p2.y,
                    p2.x, p2.y
                );
            }

            return {
                startAngle: dx > 0 ? 0 : Math.PI,
                endAngle: dx > 0 ? 0 : Math.PI
            };
        } else {
            const points = [
                { x: x1, y: y1 },
                { x: x1, y: y1 + dy * 0.25 },
                { x: x1 + dx * 0.5, y: y1 + dy * 0.4 },
                { x: x1 + dx * 0.5, y: y1 + dy * 0.6 },
                { x: x2, y: y1 + dy * 0.75 },
                { x: x2, y: y2 }
            ];

            for (let i = 0; i < points.length - 1; i++) {
                const p1 = points[i];
                const p2 = points[i + 1];
                const cpDist = 0.3;

                ctx.bezierCurveTo(
                    p1.x, p1.y + (p2.y - p1.y) * cpDist,
                    p2.x, p2.y - (p2.y - p1.y) * cpDist,
                    p2.x, p2.y
                );
            }

            return {
                startAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2,
                endAngle: dy > 0 ? Math.PI / 2 : -Math.PI / 2
            };
        }
    }
};

/**
 * Draw connectors (with arrows)
 * @param {CanvasRenderingContext2D} ctx - Canvas Context
 * @param {string} connectorType - Connection line type
 * @param {number} x1 - Start x
 * @param {number} y1 - Starting at y
 * @param {number} x2 - End x
 * @param {number} y2 - End y
 * @param {Object} style - Style Configuration
 * @ returns {Object} {startAngle, endAngle} - Angle of start and end, used to draw arrows
 */
export function drawConnector(ctx, connectorType, x1, y1, x2, y2, style = {}) {
    const pathFn = connectorPaths[connectorType];

    let angles;

    // 绘制连接线路径
    ctx.beginPath();

    if (!pathFn) {
        // 默认回退到直线
        angles = connectorPaths.straightConnector1(ctx, x1, y1, x2, y2);
    } else {
        angles = pathFn(ctx, x1, y1, x2, y2);
    }

    // 应用样式
    ctx.strokeStyle = style.stroke || '#000000';
    ctx.lineWidth = style.strokeWidth || 1;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';

    // 虚线样式
    if (style.dashStyle && style.dashStyle !== 'solid') {
        const dashPatterns = {
            dash: [5, 5],
            dashDot: [5, 5, 1, 5],
            dot: [2, 3],
            lgDash: [10, 5],
            lgDashDot: [10, 5, 2, 5],
            lgDashDotDot: [10, 5, 2, 5, 2, 5]
        };
        ctx.setLineDash(dashPatterns[style.dashStyle] || [5, 5]);
    } else {
        ctx.setLineDash([]);
    }

    ctx.stroke();

    // 重置虚线
    ctx.setLineDash([]);

    return angles || { startAngle: 0, endAngle: 0 };
}

/**
 * Draw arrows
 * @param {CanvasRenderingContext2D} ctx
 * @param {number} x - Arrow position x
 * @param {number} y - Arrow position y
 * @param {number} angle - Arrow Angle (in radians)
 * @param {number} size - Arrow Size
 * @param {string} type - Arrow Type
 * @param {string} color - Arrow Color
 */
export function drawArrowHead(ctx, x, y, angle, size, type, color) {
    ctx.save();
    ctx.translate(x, y);
    ctx.rotate(angle);
    ctx.fillStyle = color || '#000000';
    ctx.strokeStyle = color || '#000000';

    ctx.beginPath();

    if (type === 'triangle' || type === 'arrow') {
        // 三角形箭头
        ctx.moveTo(0, 0);
        ctx.lineTo(-size, -size / 2);
        ctx.lineTo(-size, size / 2);
        ctx.closePath();
        ctx.fill();
    } else if (type === 'oval') {
        // 圆形箭头
        ctx.arc(-size / 2, 0, size / 2, 0, Math.PI * 2);
        ctx.fill();
    } else if (type === 'diamond') {
        // 菱形箭头
        ctx.moveTo(0, 0);
        ctx.lineTo(-size / 2, -size / 2);
        ctx.lineTo(-size, 0);
        ctx.lineTo(-size / 2, size / 2);
        ctx.closePath();
        ctx.fill();
    } else if (type === 'stealth') {
        // 尖头箭头
        ctx.moveTo(0, 0);
        ctx.lineTo(-size, -size / 3);
        ctx.lineTo(-size * 0.7, 0);
        ctx.lineTo(-size, size / 3);
        ctx.closePath();
        ctx.fill();
    }

    ctx.restore();
}
