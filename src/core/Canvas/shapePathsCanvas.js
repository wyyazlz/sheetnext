// 形状路径绘制函数集（高性能 Canvas 实现）
// 每个函数签名：(ctx, x, y, w, h) => void

// ==================== 辅助函数（复用逻辑）====================

// 绘制正多边形
function drawPolygon(ctx, x, y, w, h, sides, rotation = -Math.PI / 2) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const radiusX = w / 2;
    const radiusY = h / 2;

    for (let i = 0; i < sides; i++) {
        const angle = (Math.PI * 2 / sides) * i + rotation;
        const px = cx + radiusX * Math.cos(angle);
        const py = cy + radiusY * Math.sin(angle);
        if (i === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
    }
    ctx.closePath();
}

// 绘制星形（通用）
function drawStar(ctx, x, y, w, h, points, innerRadiusRatio = 0.38) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const outerRadiusX = w / 2;
    const outerRadiusY = h / 2;
    const innerRadiusX = outerRadiusX * innerRadiusRatio;
    const innerRadiusY = outerRadiusY * innerRadiusRatio;

    for (let i = 0; i < points * 2; i++) {
        const angle = (Math.PI / points) * i - Math.PI / 2;
        const isOuter = i % 2 === 0;
        const rx = isOuter ? outerRadiusX : innerRadiusX;
        const ry = isOuter ? outerRadiusY : innerRadiusY;
        const px = cx + rx * Math.cos(angle);
        const py = cy + ry * Math.sin(angle);
        if (i === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
    }
    ctx.closePath();
}

// 绘制径向不规则形（爆炸形等）
function drawRadialShape(ctx, x, y, w, h, radii, rotation = -Math.PI / 2) {
    const cx = x + w / 2;
    const cy = y + h / 2;
    const radiusX = w / 2;
    const radiusY = h / 2;
    radii.forEach((ratio, index) => {
        const angle = (Math.PI * 2 / radii.length) * index + rotation;
        const px = cx + radiusX * ratio * Math.cos(angle);
        const py = cy + radiusY * ratio * Math.sin(angle);
        if (index === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
    });
    ctx.closePath();
}

// 绘制圆角矩形（通用）
function drawRoundRect(ctx, x, y, w, h, radiusRatio = 0.1) {
    const radius = Math.min(w, h) * radiusRatio;
    if (ctx.roundRect) {
        ctx.roundRect(x, y, w, h, radius);
    } else {
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + w - radius, y);
        ctx.arcTo(x + w, y, x + w, y + radius, radius);
        ctx.lineTo(x + w, y + h - radius);
        ctx.arcTo(x + w, y + h, x + w - radius, y + h, radius);
        ctx.lineTo(x + radius, y + h);
        ctx.arcTo(x, y + h, x, y + h - radius, radius);
        ctx.lineTo(x, y + radius);
        ctx.arcTo(x, y, x + radius, y, radius);
        ctx.closePath();
    }
}

// 绘制归一化路径，复杂形状统一使用 0..1 坐标
function ux(x, w, value) {
    return x + w * value;
}

function uy(y, h, value) {
    return y + h * value;
}

function drawUnitPath(ctx, x, y, w, h, commands) {
    commands.forEach((command) => {
        const [type, ...args] = command;
        switch (type) {
            case 'M':
                ctx.moveTo(ux(x, w, args[0]), uy(y, h, args[1]));
                break;
            case 'L':
                ctx.lineTo(ux(x, w, args[0]), uy(y, h, args[1]));
                break;
            case 'C':
                ctx.bezierCurveTo(
                    ux(x, w, args[0]), uy(y, h, args[1]),
                    ux(x, w, args[2]), uy(y, h, args[3]),
                    ux(x, w, args[4]), uy(y, h, args[5])
                );
                break;
            case 'Q':
                ctx.quadraticCurveTo(
                    ux(x, w, args[0]), uy(y, h, args[1]),
                    ux(x, w, args[2]), uy(y, h, args[3])
                );
                break;
            case 'Z':
                ctx.closePath();
                break;
        }
    });
}

function drawUnitPolygon(ctx, x, y, w, h, points) {
    points.forEach(([px, py], index) => {
        if (index === 0) ctx.moveTo(ux(x, w, px), uy(y, h, py));
        else ctx.lineTo(ux(x, w, px), uy(y, h, py));
    });
    ctx.closePath();
}

function drawRibbonTail(ctx, x, y, w, h, side, top, bottom, attachX) {
    const outerX = side === 'left' ? 0 : 1;
    const notchX = side === 'left' ? 0.13 : 0.87;
    const middle = (top + bottom) / 2;
    drawUnitPolygon(ctx, x, y, w, h, [
        [attachX, top],
        [outerX, top],
        [notchX, middle],
        [outerX, bottom],
        [attachX, bottom]
    ]);
}

// 绘制箭头（通用方向）
function drawArrow(ctx, x, y, w, h, direction) {
    const bodyRatio = 0.6;
    const headRatio = 0.3;

    if (direction === 'right') {
        const arrowWidth = w * (1 - headRatio);
        const arrowHeight = h * bodyRatio;
        const cy = y + h / 2;
        ctx.moveTo(x, cy - arrowHeight / 2);
        ctx.lineTo(x + arrowWidth, cy - arrowHeight / 2);
        ctx.lineTo(x + arrowWidth, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(x + arrowWidth, y + h);
        ctx.lineTo(x + arrowWidth, cy + arrowHeight / 2);
        ctx.lineTo(x, cy + arrowHeight / 2);
    } else if (direction === 'left') {
        const arrowWidth = w * headRatio;
        const arrowHeight = h * bodyRatio;
        const cy = y + h / 2;
        ctx.moveTo(x, cy);
        ctx.lineTo(x + arrowWidth, y);
        ctx.lineTo(x + arrowWidth, cy - arrowHeight / 2);
        ctx.lineTo(x + w, cy - arrowHeight / 2);
        ctx.lineTo(x + w, cy + arrowHeight / 2);
        ctx.lineTo(x + arrowWidth, cy + arrowHeight / 2);
        ctx.lineTo(x + arrowWidth, y + h);
    } else if (direction === 'up') {
        const arrowHeight = h * headRatio;
        const arrowWidth = w * bodyRatio;
        const cx = x + w / 2;
        ctx.moveTo(cx, y);
        ctx.lineTo(x + w, y + arrowHeight);
        ctx.lineTo(cx + arrowWidth / 2, y + arrowHeight);
        ctx.lineTo(cx + arrowWidth / 2, y + h);
        ctx.lineTo(cx - arrowWidth / 2, y + h);
        ctx.lineTo(cx - arrowWidth / 2, y + arrowHeight);
        ctx.lineTo(x, y + arrowHeight);
    } else if (direction === 'down') {
        const arrowHeight = h * (1 - headRatio);
        const arrowWidth = w * bodyRatio;
        const cx = x + w / 2;
        ctx.moveTo(cx - arrowWidth / 2, y);
        ctx.lineTo(cx + arrowWidth / 2, y);
        ctx.lineTo(cx + arrowWidth / 2, y + arrowHeight);
        ctx.lineTo(x + w, y + arrowHeight);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(x, y + arrowHeight);
        ctx.lineTo(cx - arrowWidth / 2, y + arrowHeight);
    }
    ctx.closePath();
}

// ==================== 形状路径映射表 ====================

export const shapePaths = {
    // ==================== 线条/连接线系列（Excel标准4种）====================

    // 直线（对角线，从左上到右下）
    line(ctx, x, y, w, h) {
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y + h);
    },

    // 直接连接符（同对角线）
    straightConnector1(ctx, x, y, w, h) {
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y + h);
    },

    // 肘形连接符
    bentConnector3(ctx, x, y, w, h) {
        const midX = x + w / 2;
        ctx.moveTo(x, y);
        ctx.lineTo(midX, y);
        ctx.lineTo(midX, y + h);
        ctx.lineTo(x + w, y + h);
    },

    // 曲线连接符
    curvedConnector3(ctx, x, y, w, h) {
        const midX = x + w / 2;
        ctx.moveTo(x, y);
        ctx.bezierCurveTo(midX, y, midX, y + h, x + w, y + h);
    },

    // ==================== 矩形系列 ====================

    rect(ctx, x, y, w, h) {
        ctx.rect(x, y, w, h);
    },

    roundRect(ctx, x, y, w, h) {
        drawRoundRect(ctx, x, y, w, h, 0.1);
    },

    snip1Rect(ctx, x, y, w, h) {
        const snip = Math.min(w, h) * 0.15;
        ctx.moveTo(x + w - snip, y);
        ctx.lineTo(x + w, y + snip);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x, y);
        ctx.closePath();
    },

    snip2SameRect(ctx, x, y, w, h) {
        const snip = Math.min(w, h) * 0.15;
        ctx.moveTo(x + snip, y);
        ctx.lineTo(x + w - snip, y);
        ctx.lineTo(x + w, y + snip);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x, y + snip);
        ctx.closePath();
    },

    snip2DiagRect(ctx, x, y, w, h) {
        const snip = Math.min(w, h) * 0.15;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w - snip, y);
        ctx.lineTo(x + w, y + snip);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x + snip, y + h);
        ctx.lineTo(x, y + h - snip);
        ctx.closePath();
    },

    snipRoundRect(ctx, x, y, w, h) {
        const snip = Math.min(w, h) * 0.15;
        const radius = Math.min(w, h) * 0.1;
        ctx.moveTo(x + w - snip, y);
        ctx.lineTo(x + w, y + snip);
        ctx.lineTo(x + w, y + h - radius);
        ctx.arcTo(x + w, y + h, x + w - radius, y + h, radius);
        ctx.lineTo(x + radius, y + h);
        ctx.arcTo(x, y + h, x, y + h - radius, radius);
        ctx.lineTo(x, y + radius);
        ctx.arcTo(x, y, x + radius, y, radius);
        ctx.lineTo(x + w - snip, y);
    },

    round1Rect(ctx, x, y, w, h) {
        const radius = Math.min(w, h) * 0.15;
        ctx.moveTo(x + w - radius, y);
        ctx.arcTo(x + w, y, x + w, y + radius, radius);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x, y);
        ctx.closePath();
    },

    round2SameRect(ctx, x, y, w, h) {
        const radius = Math.min(w, h) * 0.15;
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + w - radius, y);
        ctx.arcTo(x + w, y, x + w, y + radius, radius);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x, y + radius);
        ctx.arcTo(x, y, x + radius, y, radius);
        ctx.closePath();
    },

    round2DiagRect(ctx, x, y, w, h) {
        const radius = Math.min(w, h) * 0.15;
        ctx.moveTo(x, y + radius);
        ctx.arcTo(x, y, x + radius, y, radius);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h - radius);
        ctx.arcTo(x + w, y + h, x + w - radius, y + h, radius);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    // ==================== 基本形状 ====================

    ellipse(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
    },

    triangle(ctx, x, y, w, h) {
        ctx.moveTo(x + w / 2, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    rtTriangle(ctx, x, y, w, h) {
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    parallelogram(ctx, x, y, w, h) {
        const offset = w * 0.2;
        ctx.moveTo(x + offset, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w - offset, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    trapezoid(ctx, x, y, w, h) {
        const offset = w * 0.2;
        ctx.moveTo(x + offset, y);
        ctx.lineTo(x + w - offset, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    diamond(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        ctx.moveTo(cx, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(x, cy);
        ctx.closePath();
    },

    pentagon(ctx, x, y, w, h) {
        drawPolygon(ctx, x, y, w, h, 5);
    },

    hexagon(ctx, x, y, w, h) {
        drawPolygon(ctx, x, y, w, h, 6, 0);  // 旋转0度，底部平
    },

    heptagon(ctx, x, y, w, h) {
        drawPolygon(ctx, x, y, w, h, 7);
    },

    octagon(ctx, x, y, w, h) {
        drawPolygon(ctx, x, y, w, h, 8, Math.PI / 8);  // 平底（旋转22.5度）
    },

    decagon(ctx, x, y, w, h) {
        drawPolygon(ctx, x, y, w, h, 10, 0);  // 平底
    },

    dodecagon(ctx, x, y, w, h) {
        drawPolygon(ctx, x, y, w, h, 12, Math.PI / 12);  // 平底（旋转15度）
    },

    pie(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, Math.min(w, h) / 2, 0, Math.PI * 1.5);
        ctx.closePath();
    },

    chord(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const radius = Math.min(w, h) / 2;
        ctx.arc(cx, cy, radius, 0, Math.PI * 1.5);
        ctx.closePath();
    },

    teardrop(ctx, x, y, w, h) {
        // 标准泪滴：圆形主体 + 右上角尖角平滑过渡
        const r = Math.min(w, h) * 0.4;
        const cx = x + r + (w - r * 2) * 0.3;  // 圆心偏左
        const cy = y + h - r;  // 圆心靠下

        // 从右上角尖点开始，顺时针绘制
        ctx.moveTo(x + w, y);  // 右上角尖点
        ctx.bezierCurveTo(
            cx + r * 0.5, y,           // 控制点1
            cx + r, cy - r * 0.5,      // 控制点2
            cx + r, cy                  // 圆的右边
        );
        ctx.arc(cx, cy, r, 0, Math.PI, false);  // 下半圆
        ctx.bezierCurveTo(
            cx - r, cy - r * 0.5,      // 控制点1
            cx - r * 0.3, y,           // 控制点2
            x + w, y                    // 回到尖点
        );
        ctx.closePath();
    },

    frame(ctx, x, y, w, h) {
        const thickness = Math.min(w, h) * 0.15;
        ctx.rect(x, y, w, h);
        ctx.moveTo(x + thickness, y + thickness);
        ctx.lineTo(x + thickness, y + h - thickness);
        ctx.lineTo(x + w - thickness, y + h - thickness);
        ctx.lineTo(x + w - thickness, y + thickness);
        ctx.closePath();
    },

    halfFrame(ctx, x, y, w, h) {
        // L形半框：上边+左边，开口在右下
        const t = Math.min(w, h) * 0.2;  // 边框厚度

        // 按SVG顺序绘制（从右上内角开始）
        ctx.moveTo(x + w - t, y + t);    // 右上内角
        ctx.lineTo(x + w, y);            // 右上外角（斜切）
        ctx.lineTo(x, y);                // 左上外角
        ctx.lineTo(x, y + h);            // 左下外角
        ctx.lineTo(x + t, y + h - t);    // 左下内角（斜切）
        ctx.lineTo(x + t, y + t);        // 左上内角
        ctx.closePath();
    },

    corner(ctx, x, y, w, h) {
        // L形：下边+左边，开口在右上
        const size = Math.min(w, h) * 0.4;
        ctx.moveTo(x, y);                        // 左上
        ctx.lineTo(x + size, y);                 // 上边往右
        ctx.lineTo(x + size, y + h - size);      // 内部转角
        ctx.lineTo(x + w, y + h - size);         // 右边往右
        ctx.lineTo(x + w, y + h);                // 右下
        ctx.lineTo(x, y + h);                    // 左下
        ctx.closePath();
    },

    diagStripe(ctx, x, y, w, h) {
        const stripe = Math.min(w, h) * 0.25;
        ctx.moveTo(x + stripe, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w - stripe, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    plus(ctx, x, y, w, h) {
        const armWidth = Math.min(w, h) * 0.3;
        const cx = x + w / 2, cy = y + h / 2;
        ctx.moveTo(cx - armWidth / 2, y);
        ctx.lineTo(cx + armWidth / 2, y);
        ctx.lineTo(cx + armWidth / 2, cy - armWidth / 2);
        ctx.lineTo(x + w, cy - armWidth / 2);
        ctx.lineTo(x + w, cy + armWidth / 2);
        ctx.lineTo(cx + armWidth / 2, cy + armWidth / 2);
        ctx.lineTo(cx + armWidth / 2, y + h);
        ctx.lineTo(cx - armWidth / 2, y + h);
        ctx.lineTo(cx - armWidth / 2, cy + armWidth / 2);
        ctx.lineTo(x, cy + armWidth / 2);
        ctx.lineTo(x, cy - armWidth / 2);
        ctx.lineTo(cx - armWidth / 2, cy - armWidth / 2);
        ctx.closePath();
    },

    plaque(ctx, x, y, w, h) {
        // 牌匾形：四角内凹圆弧
        const r = Math.min(w, h) * 0.15;
        ctx.moveTo(x, y + r);
        ctx.arcTo(x + r, y + r, x + r, y, r);
        ctx.lineTo(x + w - r, y);
        ctx.arcTo(x + w - r, y + r, x + w, y + r, r);
        ctx.lineTo(x + w, y + h - r);
        ctx.arcTo(x + w - r, y + h - r, x + w - r, y + h, r);
        ctx.lineTo(x + r, y + h);
        ctx.arcTo(x + r, y + h - r, x, y + h - r, r);
        ctx.closePath();
    },

    can(ctx, x, y, w, h) {
        const topHeight = h * 0.15;
        const cx = x + w / 2;
        const radiusX = w / 2;
        const radiusY = topHeight / 2;

        // 顶部完整椭圆
        ctx.ellipse(cx, y + radiusY, radiusX, radiusY, 0, 0, Math.PI * 2);

        // 圆柱体侧面和底部（形成一个连续路径）
        ctx.moveTo(x + w, y + radiusY);
        ctx.lineTo(x + w, y + h - radiusY);
        ctx.ellipse(cx, y + h - radiusY, radiusX, radiusY, 0, 0, Math.PI, false);
        ctx.lineTo(x, y + radiusY);
    },

    cube(ctx, x, y, w, h) {
        const depth = Math.min(w, h) * 0.3;
        ctx.rect(x, y + depth, w - depth, h - depth);
        ctx.moveTo(x, y + depth);
        ctx.lineTo(x + depth, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w - depth, y + depth);
        ctx.moveTo(x + w - depth, y + depth);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h - depth);
        ctx.lineTo(x + w - depth, y + h);
    },

    bevel(ctx, x, y, w, h) {
        const bevel = Math.min(w, h) * 0.15;

        // 外部轮廓
        ctx.moveTo(x + bevel, y);
        ctx.lineTo(x + w - bevel, y);
        ctx.lineTo(x + w, y + bevel);
        ctx.lineTo(x + w, y + h - bevel);
        ctx.lineTo(x + w - bevel, y + h);
        ctx.lineTo(x + bevel, y + h);
        ctx.lineTo(x, y + h - bevel);
        ctx.lineTo(x, y + bevel);
        ctx.closePath();

        // 内部斜角线（形成立体效果）
        ctx.moveTo(x + bevel, y);
        ctx.lineTo(x + bevel, y + bevel);
        ctx.lineTo(x + w - bevel, y + bevel);
        ctx.lineTo(x + w - bevel, y);

        ctx.moveTo(x + w - bevel, y + bevel);
        ctx.lineTo(x + w - bevel, y + h - bevel);
        ctx.lineTo(x + w, y + h - bevel);

        ctx.moveTo(x + w - bevel, y + h - bevel);
        ctx.lineTo(x + bevel, y + h - bevel);
        ctx.lineTo(x + bevel, y + h);

        ctx.moveTo(x + bevel, y + h - bevel);
        ctx.lineTo(x + bevel, y + bevel);
        ctx.lineTo(x, y + bevel);
    },

    donut(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const outerR = Math.min(w, h) / 2;
        const innerR = outerR * 0.6;
        ctx.arc(cx, cy, outerR, 0, Math.PI * 2);
        ctx.moveTo(cx + innerR, cy);
        ctx.arc(cx, cy, innerR, 0, Math.PI * 2, true);
    },

    noSmoking(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const outerR = Math.min(w, h) / 2;
        const innerR = outerR * 0.7;
        ctx.arc(cx, cy, outerR, 0, Math.PI * 2);
        ctx.moveTo(cx + innerR, cy);
        ctx.arc(cx, cy, innerR, 0, Math.PI * 2, true);
        const offset = outerR * 0.707;
        ctx.moveTo(cx - offset, cy - offset);
        ctx.lineTo(cx + offset, cy + offset);
    },

    blockArc(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h;
        const outerR = Math.min(w, h);
        const innerR = outerR * 0.5;
        ctx.arc(cx, cy, outerR, Math.PI, 0);
        ctx.arc(cx, cy, innerR, 0, Math.PI, true);
        ctx.closePath();
    },

    foldedCorner(ctx, x, y, w, h) {
        // 折角形：右下角折叠（Excel标准）
        const fold = Math.min(w, h) * 0.2;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h - fold);
        ctx.lineTo(x + w - fold, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
        // 折角三角形
        ctx.moveTo(x + w, y + h - fold);
        ctx.lineTo(x + w - fold, y + h - fold);
        ctx.lineTo(x + w - fold, y + h);
    },

    smileyFace(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const r = Math.min(w, h) / 2;
        ctx.arc(cx, cy, r, 0, Math.PI * 2);
        ctx.moveTo(cx - r * 0.3 + r * 0.1, cy - r * 0.2);
        ctx.arc(cx - r * 0.3, cy - r * 0.2, r * 0.1, 0, Math.PI * 2);
        ctx.moveTo(cx + r * 0.3 + r * 0.1, cy - r * 0.2);
        ctx.arc(cx + r * 0.3, cy - r * 0.2, r * 0.1, 0, Math.PI * 2);
        ctx.moveTo(cx - r * 0.4, cy + r * 0.1);
        ctx.quadraticCurveTo(cx, cy + r * 0.5, cx + r * 0.4, cy + r * 0.1);
    },

    heart(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const topY = y + h * 0.3;
        ctx.moveTo(cx, y + h);
        ctx.bezierCurveTo(cx, y + h * 0.7, x, topY, x + w * 0.25, y);
        ctx.bezierCurveTo(x + w / 2, y, cx, topY, cx, topY);
        ctx.bezierCurveTo(cx, topY, x + w / 2, y, x + w * 0.75, y);
        ctx.bezierCurveTo(x + w, topY, cx, y + h * 0.7, cx, y + h);
    },

    lightningBolt(ctx, x, y, w, h) {
        drawUnitPolygon(ctx, x, y, w, h, [
            [0.3909, 0],
            [0, 0.1527],
            [0.3878, 0.3837],
            [0.2726, 0.5038],
            [0.6164, 0.6829],
            [0.5037, 0.859],
            [1, 1],
            [0.7715, 0.7618],
            [0.8272, 0.6458],
            [0.5319, 0.4612],
            [0.6526, 0.2726]
        ]);
    },

    sun(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const innerRadius = Math.min(w, h) * 0.3;
        const outerRadius = Math.min(w, h) * 0.5;
        const rays = 12;
        for (let i = 0; i < rays * 2; i++) {
            const angle = (Math.PI / rays) * i;
            const radius = i % 2 === 0 ? outerRadius : innerRadius;
            const px = cx + radius * Math.cos(angle);
            const py = cy + radius * Math.sin(angle);
            if (i === 0) ctx.moveTo(px, py);
            else ctx.lineTo(px, py);
        }
        ctx.closePath();
    },

    moon(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const r = Math.min(w, h) / 2;
        const offset = r * 0.3;
        ctx.arc(cx, cy, r, 0, Math.PI * 2, false);
        ctx.arc(cx + offset, cy, r * 0.8, 0, Math.PI * 2, true);
    },

    cloud(ctx, x, y, w, h) {
        drawUnitPath(ctx, x, y, w, h, [
            ['M', 0.18, 0.74],
            ['C', 0.08, 0.74, 0, 0.65, 0, 0.54],
            ['C', 0, 0.44, 0.08, 0.36, 0.2, 0.34],
            ['C', 0.25, 0.16, 0.42, 0.07, 0.58, 0.14],
            ['C', 0.71, 0.18, 0.79, 0.29, 0.82, 0.42],
            ['C', 0.92, 0.41, 1, 0.5, 1, 0.61],
            ['C', 1, 0.74, 0.89, 0.83, 0.77, 0.8],
            ['C', 0.72, 0.92, 0.57, 0.96, 0.47, 0.85],
            ['C', 0.36, 0.94, 0.22, 0.88, 0.18, 0.74],
            ['Z']
        ]);
    },

    arc(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const radius = Math.min(w, h) / 2;
        ctx.arc(cx, cy, radius, Math.PI, 0);
    },

    // ==================== 括号系列 ====================

    bracketPair(ctx, x, y, w, h) {
        const thickness = w * 0.15;
        const depth = w * 0.1;
        ctx.moveTo(x + thickness, y);
        ctx.lineTo(x, y);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x + thickness, y + h);
        ctx.moveTo(x + w - thickness, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x + w - thickness, y + h);
    },

    bracePair(ctx, x, y, w, h) {
        // 大括号对：中间尖点不超出边界
        const inset = w * 0.15;  // 内缩距离
        const curveH = h * 0.2;  // 曲线高度
        const cy = y + h / 2;
        // 左括号 {
        ctx.moveTo(x + inset, y);
        ctx.quadraticCurveTo(x + inset * 0.3, y, x + inset * 0.3, y + curveH);
        ctx.lineTo(x + inset * 0.3, cy - curveH * 0.5);
        ctx.quadraticCurveTo(x + inset * 0.3, cy, x, cy);  // 尖点在x处
        ctx.quadraticCurveTo(x + inset * 0.3, cy, x + inset * 0.3, cy + curveH * 0.5);
        ctx.lineTo(x + inset * 0.3, y + h - curveH);
        ctx.quadraticCurveTo(x + inset * 0.3, y + h, x + inset, y + h);
        // 右括号 }
        ctx.moveTo(x + w - inset, y);
        ctx.quadraticCurveTo(x + w - inset * 0.3, y, x + w - inset * 0.3, y + curveH);
        ctx.lineTo(x + w - inset * 0.3, cy - curveH * 0.5);
        ctx.quadraticCurveTo(x + w - inset * 0.3, cy, x + w, cy);  // 尖点在x+w处
        ctx.quadraticCurveTo(x + w - inset * 0.3, cy, x + w - inset * 0.3, cy + curveH * 0.5);
        ctx.lineTo(x + w - inset * 0.3, y + h - curveH);
        ctx.quadraticCurveTo(x + w - inset * 0.3, y + h, x + w - inset, y + h);
    },

    leftBracket(ctx, x, y, w, h) {
        const thickness = w * 0.3;
        ctx.moveTo(x + thickness, y);
        ctx.lineTo(x, y);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x + thickness, y + h);
    },

    rightBracket(ctx, x, y, w, h) {
        const thickness = w * 0.3;
        ctx.moveTo(x + w - thickness, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x + w - thickness, y + h);
    },

    leftBrace(ctx, x, y, w, h) {
        // 左大括号：中间尖点在左边界
        const inset = w * 0.4;
        const curveH = h * 0.2;
        const cy = y + h / 2;
        ctx.moveTo(x + inset, y);
        ctx.quadraticCurveTo(x + inset * 0.3, y, x + inset * 0.3, y + curveH);
        ctx.lineTo(x + inset * 0.3, cy - curveH * 0.5);
        ctx.quadraticCurveTo(x + inset * 0.3, cy, x, cy);  // 尖点在x处
        ctx.quadraticCurveTo(x + inset * 0.3, cy, x + inset * 0.3, cy + curveH * 0.5);
        ctx.lineTo(x + inset * 0.3, y + h - curveH);
        ctx.quadraticCurveTo(x + inset * 0.3, y + h, x + inset, y + h);
    },

    rightBrace(ctx, x, y, w, h) {
        // 右大括号：中间尖点在右边界
        const inset = w * 0.4;
        const curveH = h * 0.2;
        const cy = y + h / 2;
        ctx.moveTo(x + w - inset, y);
        ctx.quadraticCurveTo(x + w - inset * 0.3, y, x + w - inset * 0.3, y + curveH);
        ctx.lineTo(x + w - inset * 0.3, cy - curveH * 0.5);
        ctx.quadraticCurveTo(x + w - inset * 0.3, cy, x + w, cy);  // 尖点在x+w处
        ctx.quadraticCurveTo(x + w - inset * 0.3, cy, x + w - inset * 0.3, cy + curveH * 0.5);
        ctx.lineTo(x + w - inset * 0.3, y + h - curveH);
        ctx.quadraticCurveTo(x + w - inset * 0.3, y + h, x + w - inset, y + h);
    },

    // ==================== 箭头系列 ====================

    rightArrow(ctx, x, y, w, h) {
        drawArrow(ctx, x, y, w, h, 'right');
    },

    leftArrow(ctx, x, y, w, h) {
        drawArrow(ctx, x, y, w, h, 'left');
    },

    upArrow(ctx, x, y, w, h) {
        drawArrow(ctx, x, y, w, h, 'up');
    },

    downArrow(ctx, x, y, w, h) {
        drawArrow(ctx, x, y, w, h, 'down');
    },

    leftRightArrow(ctx, x, y, w, h) {
        const arrowHeight = h * 0.6;
        const cy = y + h / 2;
        const headWidth = w * 0.2;
        ctx.moveTo(x, cy);
        ctx.lineTo(x + headWidth, y);
        ctx.lineTo(x + headWidth, cy - arrowHeight / 2);
        ctx.lineTo(x + w - headWidth, cy - arrowHeight / 2);
        ctx.lineTo(x + w - headWidth, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(x + w - headWidth, y + h);
        ctx.lineTo(x + w - headWidth, cy + arrowHeight / 2);
        ctx.lineTo(x + headWidth, cy + arrowHeight / 2);
        ctx.lineTo(x + headWidth, y + h);
        ctx.closePath();
    },

    upDownArrow(ctx, x, y, w, h) {
        const arrowWidth = w * 0.6;
        const cx = x + w / 2;
        const headHeight = h * 0.2;
        ctx.moveTo(cx, y);
        ctx.lineTo(x + w, y + headHeight);
        ctx.lineTo(cx + arrowWidth / 2, y + headHeight);
        ctx.lineTo(cx + arrowWidth / 2, y + h - headHeight);
        ctx.lineTo(x + w, y + h - headHeight);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(x, y + h - headHeight);
        ctx.lineTo(cx - arrowWidth / 2, y + h - headHeight);
        ctx.lineTo(cx - arrowWidth / 2, y + headHeight);
        ctx.lineTo(x, y + headHeight);
        ctx.closePath();
    },

    quadArrow(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        const arrowSize = Math.min(w, h) * 0.35;
        const headSize = arrowSize * 0.4;
        ctx.moveTo(cx, y);
        ctx.lineTo(cx + headSize, y + headSize);
        ctx.lineTo(cx + arrowSize / 2, y + headSize);
        ctx.lineTo(cx + arrowSize / 2, cy - arrowSize / 2);
        ctx.lineTo(cx + headSize, cy - arrowSize / 2);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(x + w - headSize, cy + arrowSize / 2);
        ctx.lineTo(cx + arrowSize / 2, cy + arrowSize / 2);
        ctx.lineTo(cx + arrowSize / 2, y + h - headSize);
        ctx.lineTo(cx + headSize, y + h - headSize);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(cx - headSize, y + h - headSize);
        ctx.lineTo(cx - arrowSize / 2, y + h - headSize);
        ctx.lineTo(cx - arrowSize / 2, cy + arrowSize / 2);
        ctx.lineTo(x + headSize, cy + arrowSize / 2);
        ctx.lineTo(x, cy);
        ctx.lineTo(x + headSize, cy - arrowSize / 2);
        ctx.lineTo(cx - arrowSize / 2, cy - arrowSize / 2);
        ctx.lineTo(cx - arrowSize / 2, y + headSize);
        ctx.lineTo(cx - headSize, y + headSize);
        ctx.closePath();
    },

    leftRightUpArrow(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const bodyWidth = w * 0.3;
        const bodyHeight = h * 0.5;
        const headSize = w * 0.2;
        ctx.moveTo(cx, y);
        ctx.lineTo(cx + headSize, y + headSize);
        ctx.lineTo(cx + bodyWidth / 2, y + headSize);
        ctx.lineTo(cx + bodyWidth / 2, y + bodyHeight);
        ctx.lineTo(x + w - headSize, y + bodyHeight);
        ctx.lineTo(x + w - headSize, y + bodyHeight - bodyWidth / 2);
        ctx.lineTo(x + w, y + bodyHeight + headSize);
        ctx.lineTo(x + w - headSize, y + bodyHeight + headSize * 2);
        ctx.lineTo(x + w - headSize, y + bodyHeight + bodyWidth / 2);
        ctx.lineTo(cx + bodyWidth / 2, y + bodyHeight + bodyWidth / 2);
        ctx.lineTo(cx + bodyWidth / 2, y + h);
        ctx.lineTo(cx - bodyWidth / 2, y + h);
        ctx.lineTo(cx - bodyWidth / 2, y + bodyHeight + bodyWidth / 2);
        ctx.lineTo(x + headSize, y + bodyHeight + bodyWidth / 2);
        ctx.lineTo(x + headSize, y + bodyHeight + headSize * 2);
        ctx.lineTo(x, y + bodyHeight + headSize);
        ctx.lineTo(x + headSize, y + bodyHeight - bodyWidth / 2);
        ctx.lineTo(x + headSize, y + bodyHeight);
        ctx.lineTo(cx - bodyWidth / 2, y + bodyHeight);
        ctx.lineTo(cx - bodyWidth / 2, y + headSize);
        ctx.lineTo(cx - headSize, y + headSize);
        ctx.closePath();
    },

    bentArrow(ctx, x, y, w, h) {
        const bodyWidth = w * 0.3;
        const headHeight = h * 0.4;
        ctx.moveTo(x, y + h);
        ctx.lineTo(x, y + headHeight);
        ctx.lineTo(x + w - headHeight, y + headHeight);
        ctx.lineTo(x + w - headHeight, y);
        ctx.lineTo(x + w, y + headHeight * 1.5);
        ctx.lineTo(x + w - headHeight, y + headHeight * 3);
        ctx.lineTo(x + w - headHeight, y + headHeight * 2);
        ctx.lineTo(x + bodyWidth, y + headHeight * 2);
        ctx.lineTo(x + bodyWidth, y + h);
        ctx.closePath();
    },

    uturnArrow(ctx, x, y, w, h) {
        const bodyWidth = w * 0.3;
        const headHeight = h * 0.3;
        const radius = (w - bodyWidth) / 2;
        ctx.moveTo(x, y + headHeight);
        ctx.lineTo(x + headHeight, y);
        ctx.lineTo(x + headHeight * 2, y + headHeight);
        ctx.lineTo(x + headHeight, y + headHeight);
        ctx.arc(x + headHeight + radius, y + headHeight + radius, radius, Math.PI, 0, true);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x + w - bodyWidth, y + h);
        ctx.lineTo(x + w - bodyWidth, y + headHeight + radius);
        ctx.arc(x + headHeight + radius, y + headHeight + radius, radius - bodyWidth, 0, Math.PI);
        ctx.closePath();
    },

    leftUpArrow(ctx, x, y, w, h) {
        const bodyWidth = w * 0.3;
        const headSize = w * 0.3;
        ctx.moveTo(x, y + h / 2);
        ctx.lineTo(x + headSize, y + h / 2 - headSize);
        ctx.lineTo(x + headSize, y + h / 2 - bodyWidth / 2);
        ctx.lineTo(x + w / 2 - bodyWidth / 2, y + h / 2 - bodyWidth / 2);
        ctx.lineTo(x + w / 2 - bodyWidth / 2, y + headSize);
        ctx.lineTo(x + w / 2 - headSize, y + headSize);
        ctx.lineTo(x + w / 2, y);
        ctx.lineTo(x + w / 2 + headSize, y + headSize);
        ctx.lineTo(x + w / 2 + bodyWidth / 2, y + headSize);
        ctx.lineTo(x + w / 2 + bodyWidth / 2, y + h / 2 + bodyWidth / 2);
        ctx.lineTo(x + headSize, y + h / 2 + bodyWidth / 2);
        ctx.lineTo(x + headSize, y + h / 2 + headSize);
        ctx.closePath();
    },

    bentUpArrow(ctx, x, y, w, h) {
        const bodyWidth = h * 0.3;
        const headWidth = w * 0.4;
        ctx.moveTo(x + w / 2, y);
        ctx.lineTo(x + w, y + headWidth);
        ctx.lineTo(x + w / 2 + bodyWidth / 2, y + headWidth);
        ctx.lineTo(x + w / 2 + bodyWidth / 2, y + h - bodyWidth);
        ctx.lineTo(x, y + h - bodyWidth);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x + w / 2 + bodyWidth / 2, y + h);
        ctx.lineTo(x + w / 2 + bodyWidth / 2, y + headWidth * 2);
        ctx.lineTo(x + w / 2 - bodyWidth / 2, y + headWidth * 2);
        ctx.lineTo(x + w / 2 - bodyWidth / 2, y + h);
        ctx.lineTo(x + w / 2 - bodyWidth / 2, y + headWidth);
        ctx.lineTo(x, y + headWidth);
        ctx.closePath();
    },

    curvedRightArrow(ctx, x, y, w, h) {
        const cy = y + h / 2;
        const bodyHeight = h * 0.4;
        const headHeight = h * 0.6;
        ctx.moveTo(x, cy - bodyHeight / 2);
        ctx.quadraticCurveTo(x + w * 0.5, y, x + w - headHeight, y);
        ctx.lineTo(x + w - headHeight, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(x + w - headHeight, y + h);
        ctx.lineTo(x + w - headHeight, y + h);
        ctx.quadraticCurveTo(x + w * 0.5, cy + bodyHeight / 2, x, cy + bodyHeight / 2);
        ctx.closePath();
    },

    curvedLeftArrow(ctx, x, y, w, h) {
        const cy = y + h / 2;
        const bodyHeight = h * 0.4;
        const headHeight = h * 0.6;
        ctx.moveTo(x + w, cy - bodyHeight / 2);
        ctx.quadraticCurveTo(x + w * 0.5, y, x + headHeight, y);
        ctx.lineTo(x + headHeight, y);
        ctx.lineTo(x, cy);
        ctx.lineTo(x + headHeight, y + h);
        ctx.lineTo(x + headHeight, y + h);
        ctx.quadraticCurveTo(x + w * 0.5, cy + bodyHeight / 2, x + w, cy + bodyHeight / 2);
        ctx.closePath();
    },

    curvedUpArrow(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const bodyWidth = w * 0.4;
        const headWidth = w * 0.6;
        ctx.moveTo(cx - bodyWidth / 2, y + h);
        ctx.quadraticCurveTo(x, y + h * 0.5, x, y + headWidth);
        ctx.lineTo(x, y + headWidth);
        ctx.lineTo(cx, y);
        ctx.lineTo(x + w, y + headWidth);
        ctx.lineTo(x + w, y + headWidth);
        ctx.quadraticCurveTo(cx + bodyWidth / 2, y + h * 0.5, cx + bodyWidth / 2, y + h);
        ctx.closePath();
    },

    curvedDownArrow(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const bodyWidth = w * 0.4;
        const headWidth = w * 0.6;
        ctx.moveTo(cx - bodyWidth / 2, y);
        ctx.quadraticCurveTo(x, y + h * 0.5, x, y + h - headWidth);
        ctx.lineTo(x, y + h - headWidth);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(x + w, y + h - headWidth);
        ctx.lineTo(x + w, y + h - headWidth);
        ctx.quadraticCurveTo(cx + bodyWidth / 2, y + h * 0.5, cx + bodyWidth / 2, y);
        ctx.closePath();
    },

    stripedRightArrow(ctx, x, y, w, h) {
        const bodyHeight = h * 0.6;
        const cy = y + h / 2;
        const headWidth = w * 0.3;
        const stripeWidth = w * 0.15;
        ctx.moveTo(x + stripeWidth * 2, cy - bodyHeight / 2);
        ctx.lineTo(x + w - headWidth, cy - bodyHeight / 2);
        ctx.lineTo(x + w - headWidth, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(x + w - headWidth, y + h);
        ctx.lineTo(x + w - headWidth, cy + bodyHeight / 2);
        ctx.lineTo(x + stripeWidth * 2, cy + bodyHeight / 2);
        ctx.closePath();
        ctx.moveTo(x, y + h * 0.2);
        ctx.lineTo(x, y + h * 0.8);
        ctx.moveTo(x + stripeWidth, y + h * 0.2);
        ctx.lineTo(x + stripeWidth, y + h * 0.8);
    },

    notchedRightArrow(ctx, x, y, w, h) {
        const arrowHeight = h * 0.6;
        const cy = y + h / 2;
        const notchDepth = w * 0.15;
        ctx.moveTo(x, cy - arrowHeight / 2);
        ctx.lineTo(x + w - w * 0.3, cy - arrowHeight / 2);
        ctx.lineTo(x + w - w * 0.3, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(x + w - w * 0.3, y + h);
        ctx.lineTo(x + w - w * 0.3, cy + arrowHeight / 2);
        ctx.lineTo(x, cy + arrowHeight / 2);
        ctx.lineTo(x + notchDepth, cy);
        ctx.closePath();
    },

    homePlate(ctx, x, y, w, h) {
        const notch = w * 0.3;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w - notch, y);
        ctx.lineTo(x + w, y + h / 2);
        ctx.lineTo(x + w - notch, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    chevron(ctx, x, y, w, h) {
        const depth = w * 0.3;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w - depth, y);
        ctx.lineTo(x + w, y + h / 2);
        ctx.lineTo(x + w - depth, y + h);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x + depth, y + h / 2);
        ctx.closePath();
    },

    rightArrowCallout(ctx, x, y, w, h) {
        const calloutW = w * 0.6;
        const arrowH = h * 0.4;
        const cy = y + h / 2;
        ctx.moveTo(x, y);
        ctx.lineTo(x + calloutW, y);
        ctx.lineTo(x + calloutW, cy - arrowH / 2);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(x + calloutW, cy + arrowH / 2);
        ctx.lineTo(x + calloutW, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    downArrowCallout(ctx, x, y, w, h) {
        const calloutH = h * 0.6;
        const arrowW = w * 0.4;
        const cx = x + w / 2;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + calloutH);
        ctx.lineTo(cx + arrowW / 2, y + calloutH);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(cx - arrowW / 2, y + calloutH);
        ctx.lineTo(x, y + calloutH);
        ctx.closePath();
    },

    leftArrowCallout(ctx, x, y, w, h) {
        const calloutW = w * 0.6;
        const arrowH = h * 0.4;
        const cy = y + h / 2;
        ctx.moveTo(x + w, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x + w - calloutW, y + h);
        ctx.lineTo(x + w - calloutW, cy + arrowH / 2);
        ctx.lineTo(x, cy);
        ctx.lineTo(x + w - calloutW, cy - arrowH / 2);
        ctx.lineTo(x + w - calloutW, y);
        ctx.closePath();
    },

    upArrowCallout(ctx, x, y, w, h) {
        const calloutH = h * 0.6;
        const arrowW = w * 0.4;
        const cx = x + w / 2;
        ctx.moveTo(x, y + h);
        ctx.lineTo(x, y + h - calloutH);
        ctx.lineTo(cx - arrowW / 2, y + h - calloutH);
        ctx.lineTo(cx, y);
        ctx.lineTo(cx + arrowW / 2, y + h - calloutH);
        ctx.lineTo(x + w, y + h - calloutH);
        ctx.lineTo(x + w, y + h);
        ctx.closePath();
    },

    leftRightArrowCallout(ctx, x, y, w, h) {
        const calloutW = w * 0.4;
        const arrowH = h * 0.3;
        const cy = y + h / 2;
        const calloutX1 = x + w * 0.3;
        const calloutX2 = x + w * 0.7;
        ctx.moveTo(calloutX1, y);
        ctx.lineTo(calloutX2, y);
        ctx.lineTo(calloutX2, cy - arrowH / 2);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(calloutX2, cy + arrowH / 2);
        ctx.lineTo(calloutX2, y + h);
        ctx.lineTo(calloutX1, y + h);
        ctx.lineTo(calloutX1, cy + arrowH / 2);
        ctx.lineTo(x, cy);
        ctx.lineTo(calloutX1, cy - arrowH / 2);
        ctx.closePath();
    },

    quadArrowCallout(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const calloutSize = Math.min(w, h) * 0.3;
        const arrowSize = Math.min(w, h) * 0.2;
        ctx.moveTo(cx - calloutSize, y + calloutSize);
        ctx.lineTo(cx + calloutSize, y + calloutSize);
        ctx.lineTo(cx + calloutSize, cy - arrowSize);
        ctx.lineTo(cx + arrowSize, cy - arrowSize);
        ctx.lineTo(cx + arrowSize, y);
        ctx.lineTo(cx, y - arrowSize);
        ctx.lineTo(cx - arrowSize, y);
        ctx.lineTo(cx - arrowSize, cy - arrowSize);
        ctx.lineTo(cx - calloutSize, cy - arrowSize);
        ctx.lineTo(cx - calloutSize, cy + arrowSize);
        ctx.lineTo(cx - arrowSize, cy + arrowSize);
        ctx.lineTo(cx - arrowSize, y + h);
        ctx.lineTo(cx, y + h + arrowSize);
        ctx.lineTo(cx + arrowSize, y + h);
        ctx.lineTo(cx + arrowSize, cy + arrowSize);
        ctx.lineTo(cx + calloutSize, cy + arrowSize);
        ctx.lineTo(cx + calloutSize, y + h - calloutSize);
        ctx.lineTo(x + w - calloutSize, y + h - calloutSize);
        ctx.lineTo(x + w - calloutSize, cy + arrowSize);
        ctx.lineTo(x + w, cy + arrowSize);
        ctx.lineTo(x + w + arrowSize, cy);
        ctx.lineTo(x + w, cy - arrowSize);
        ctx.lineTo(x + w - calloutSize, cy - arrowSize);
        ctx.lineTo(x + w - calloutSize, y + calloutSize);
        ctx.closePath();
    },

    circularArrow(ctx, x, y, w, h) {
        const cx = x + w / 2;
        const cy = y + h / 2;
        const radius = Math.min(w, h) * 0.4;
        const arrowSize = radius * 0.3;
        ctx.arc(cx, cy, radius, Math.PI * 0.2, Math.PI * 1.8);
        ctx.lineTo(cx - radius * 0.5, cy - radius * 0.8 - arrowSize);
        ctx.lineTo(cx - radius * 0.8, cy - radius * 0.5);
        ctx.lineTo(cx - radius * 0.5 - arrowSize, cy - radius * 0.8);
        ctx.lineTo(cx - radius * 0.5, cy - radius * 0.8 + arrowSize);
    },

    // ==================== 公式系列 ====================

    mathPlus(ctx, x, y, w, h) {
        const thickness = Math.min(w, h) * 0.15;
        const cx = x + w / 2;
        const cy = y + h / 2;
        ctx.moveTo(cx - thickness / 2, y);
        ctx.lineTo(cx + thickness / 2, y);
        ctx.lineTo(cx + thickness / 2, cy - thickness / 2);
        ctx.lineTo(x + w, cy - thickness / 2);
        ctx.lineTo(x + w, cy + thickness / 2);
        ctx.lineTo(cx + thickness / 2, cy + thickness / 2);
        ctx.lineTo(cx + thickness / 2, y + h);
        ctx.lineTo(cx - thickness / 2, y + h);
        ctx.lineTo(cx - thickness / 2, cy + thickness / 2);
        ctx.lineTo(x, cy + thickness / 2);
        ctx.lineTo(x, cy - thickness / 2);
        ctx.lineTo(cx - thickness / 2, cy - thickness / 2);
        ctx.closePath();
    },

    mathMinus(ctx, x, y, w, h) {
        const thickness = Math.min(w, h) * 0.15;
        const cy = y + h / 2;
        ctx.rect(x, cy - thickness / 2, w, thickness);
    },

    mathMultiply(ctx, x, y, w, h) {
        const thickness = Math.min(w, h) * 0.15;
        const cx = x + w / 2;
        const cy = y + h / 2;
        const offset = Math.min(w, h) * 0.35;
        ctx.save();
        ctx.translate(cx, cy);
        ctx.rotate(Math.PI / 4);
        ctx.rect(-thickness / 2, -offset, thickness, offset * 2);
        ctx.rect(-offset, -thickness / 2, offset * 2, thickness);
        ctx.restore();
    },

    mathDivide(ctx, x, y, w, h) {
        const thickness = Math.min(w, h) * 0.15;
        const cy = y + h / 2;
        const dotRadius = thickness * 1.2;
        ctx.rect(x, cy - thickness / 2, w, thickness);
        ctx.moveTo(x + w / 2 + dotRadius, y + h * 0.25);
        ctx.arc(x + w / 2, y + h * 0.25, dotRadius, 0, Math.PI * 2);
        ctx.moveTo(x + w / 2 + dotRadius, y + h * 0.75);
        ctx.arc(x + w / 2, y + h * 0.75, dotRadius, 0, Math.PI * 2);
    },

    mathEqual(ctx, x, y, w, h) {
        const thickness = Math.min(w, h) * 0.15;
        const gap = h * 0.2;
        const cy = y + h / 2;
        ctx.rect(x, cy - gap / 2 - thickness, w, thickness);
        ctx.rect(x, cy + gap / 2, w, thickness);
    },

    mathNotEqual(ctx, x, y, w, h) {
        const thickness = Math.min(w, h) * 0.15;
        const gap = h * 0.2;
        const cy = y + h / 2;
        ctx.rect(x, cy - gap / 2 - thickness, w, thickness);
        ctx.rect(x, cy + gap / 2, w, thickness);
        ctx.save();
        ctx.translate(x + w / 2, cy);
        ctx.rotate(Math.PI / 4);
        ctx.rect(-thickness / 2, -h * 0.6, thickness, h * 1.2);
        ctx.restore();
    },

    // ==================== 流程图（复用之前的实现）====================

    flowChartProcess(ctx, x, y, w, h) {
        ctx.rect(x, y, w, h);
    },

    flowChartAlternateProcess(ctx, x, y, w, h) {
        drawRoundRect(ctx, x, y, w, h, 0.15);
    },

    flowChartDecision(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        ctx.moveTo(cx, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(x, cy);
        ctx.closePath();
    },

    flowChartInputOutput(ctx, x, y, w, h) {
        const offset = w * 0.15;
        ctx.moveTo(x + offset, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w - offset, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    flowChartPredefinedProcess(ctx, x, y, w, h) {
        const inset = w * 0.1;
        ctx.rect(x, y, w, h);
        ctx.moveTo(x + inset, y);
        ctx.lineTo(x + inset, y + h);
        ctx.moveTo(x + w - inset, y);
        ctx.lineTo(x + w - inset, y + h);
    },

    flowChartInternalStorage(ctx, x, y, w, h) {
        const offsetX = w * 0.15;
        const offsetY = h * 0.15;
        ctx.rect(x, y, w, h);
        ctx.moveTo(x, y + offsetY);
        ctx.lineTo(x + w, y + offsetY);
        ctx.moveTo(x + offsetX, y);
        ctx.lineTo(x + offsetX, y + h);
    },

    flowChartDocument(ctx, x, y, w, h) {
        const waveHeight = h * 0.1;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h - waveHeight);
        ctx.quadraticCurveTo(x + w * 0.75, y + h, x + w / 2, y + h - waveHeight);
        ctx.quadraticCurveTo(x + w * 0.25, y + h - waveHeight * 2, x, y + h - waveHeight);
        ctx.closePath();
    },

    flowChartMultidocument(ctx, x, y, w, h) {
        const offset = Math.min(w, h) * 0.05;
        const waveHeight = h * 0.08;
        ctx.moveTo(x + offset * 2, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h - offset * 2 - waveHeight);
        ctx.quadraticCurveTo(x + w * 0.75, y + h - offset * 2, x + w / 2, y + h - offset * 2 - waveHeight);
        ctx.quadraticCurveTo(x + w * 0.25, y + h - offset * 2 - waveHeight * 2, x + offset * 2, y + h - offset * 2 - waveHeight);
        ctx.closePath();
        ctx.moveTo(x + offset, y + offset);
        ctx.lineTo(x + w - offset, y + offset);
        ctx.moveTo(x + w - offset, y + offset);
        ctx.lineTo(x + w - offset, y + h - offset - waveHeight);
        ctx.moveTo(x, y + offset * 2);
        ctx.lineTo(x + w - offset * 2, y + offset * 2);
        ctx.moveTo(x + w - offset * 2, y + offset * 2);
        ctx.lineTo(x + w - offset * 2, y + h - waveHeight);
    },

    flowChartTerminator(ctx, x, y, w, h) {
        const radius = h / 2;
        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + w - radius, y);
        ctx.arcTo(x + w, y, x + w, y + radius, radius);
        ctx.arcTo(x + w, y + h, x + w - radius, y + h, radius);
        ctx.lineTo(x + radius, y + h);
        ctx.arcTo(x, y + h, x, y + h - radius, radius);
        ctx.arcTo(x, y, x + radius, y, radius);
        ctx.closePath();
    },

    flowChartPreparation(ctx, x, y, w, h) {
        const offset = w * 0.15;
        ctx.moveTo(x + offset, y);
        ctx.lineTo(x + w - offset, y);
        ctx.lineTo(x + w, y + h / 2);
        ctx.lineTo(x + w - offset, y + h);
        ctx.lineTo(x + offset, y + h);
        ctx.lineTo(x, y + h / 2);
        ctx.closePath();
    },

    flowChartManualInput(ctx, x, y, w, h) {
        const offset = h * 0.15;
        ctx.moveTo(x, y + offset);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    flowChartManualOperation(ctx, x, y, w, h) {
        const offset = w * 0.1;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w - offset, y + h);
        ctx.lineTo(x + offset, y + h);
        ctx.closePath();
    },

    flowChartConnector(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        ctx.ellipse(cx, cy, w / 2, h / 2, 0, 0, Math.PI * 2);
    },

    flowChartOffpageConnector(ctx, x, y, w, h) {
        const notchHeight = h * 0.15;
        const cx = x + w / 2;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h - notchHeight);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(x, y + h - notchHeight);
        ctx.closePath();
    },

    flowChartPunchedCard(ctx, x, y, w, h) {
        const offset = w * 0.1;
        ctx.moveTo(x + offset, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.lineTo(x, y);
        ctx.closePath();
    },

    flowChartPunchedTape(ctx, x, y, w, h) {
        const waveHeight = h * 0.1;
        ctx.moveTo(x, y + waveHeight);
        ctx.quadraticCurveTo(x + w * 0.25, y, x + w / 2, y + waveHeight);
        ctx.quadraticCurveTo(x + w * 0.75, y + waveHeight * 2, x + w, y + waveHeight);
        ctx.lineTo(x + w, y + h - waveHeight);
        ctx.quadraticCurveTo(x + w * 0.75, y + h, x + w / 2, y + h - waveHeight);
        ctx.quadraticCurveTo(x + w * 0.25, y + h - waveHeight * 2, x, y + h - waveHeight);
        ctx.closePath();
    },

    flowChartSummingJunction(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        const radius = Math.min(w, h) / 2;
        ctx.ellipse(cx, cy, radius, radius, 0, 0, Math.PI * 2);
        const offset = radius * 0.707;
        ctx.moveTo(cx - offset, cy - offset);
        ctx.lineTo(cx + offset, cy + offset);
        ctx.moveTo(cx + offset, cy - offset);
        ctx.lineTo(cx - offset, cy + offset);
    },

    flowChartOr(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        const radius = Math.min(w, h) / 2;
        ctx.ellipse(cx, cy, radius, radius, 0, 0, Math.PI * 2);
        ctx.moveTo(cx, cy - radius);
        ctx.lineTo(cx, cy + radius);
        ctx.moveTo(cx - radius, cy);
        ctx.lineTo(cx + radius, cy);
    },

    flowChartCollate(ctx, x, y, w, h) {
        const cx = x + w / 2;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(cx, y + h / 2);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.lineTo(cx, y + h / 2);
        ctx.closePath();
    },

    flowChartSort(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        ctx.moveTo(cx, y);
        ctx.lineTo(x + w, cy);
        ctx.lineTo(cx, y + h);
        ctx.lineTo(x, cy);
        ctx.closePath();
        ctx.moveTo(x, cy);
        ctx.lineTo(x + w, cy);
    },

    flowChartExtract(ctx, x, y, w, h) {
        ctx.moveTo(x + w / 2, y);
        ctx.lineTo(x + w, y + h);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    flowChartMerge(ctx, x, y, w, h) {
        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w / 2, y + h);
        ctx.closePath();
    },

    flowChartOnlineStorage(ctx, x, y, w, h) {
        const arcWidth = w * 0.1;
        ctx.moveTo(x + arcWidth, y);
        ctx.lineTo(x + w, y);
        ctx.arcTo(x + w + arcWidth, y + h / 2, x + w, y + h, h / 2);
        ctx.lineTo(x + arcWidth, y + h);
        ctx.arcTo(x - arcWidth, y + h / 2, x + arcWidth, y, h / 2);
        ctx.closePath();
    },

    flowChartDelay(ctx, x, y, w, h) {
        const arcStart = w * 0.7;
        ctx.moveTo(x, y);
        ctx.lineTo(x + arcStart, y);
        ctx.arcTo(x + w, y, x + w, y + h / 2, h / 2);
        ctx.arcTo(x + w, y + h, x + arcStart, y + h, h / 2);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    flowChartMagneticTape(ctx, x, y, w, h) {
        const cx = x + w / 2, cy = y + h / 2;
        const radiusX = w / 2, radiusY = h / 2;
        ctx.ellipse(cx, cy, radiusX, radiusY, 0, Math.PI, Math.PI * 2);
        ctx.lineTo(x, y + h);
        ctx.arcTo(x, y + h + radiusY, cx, y + h + radiusY, radiusX);
        ctx.arcTo(x + w, y + h + radiusY, x + w, y + h, radiusX);
        ctx.closePath();
    },

    flowChartMagneticDisk(ctx, x, y, w, h) {
        const topHeight = h * 0.15;
        const cx = x + w / 2;
        ctx.ellipse(cx, y + topHeight / 2, w / 2, topHeight / 2, 0, 0, Math.PI * 2);
        ctx.moveTo(x, y + topHeight / 2);
        ctx.lineTo(x, y + h - topHeight / 2);
        ctx.ellipse(cx, y + h - topHeight / 2, w / 2, topHeight / 2, 0, 0, Math.PI);
        ctx.lineTo(x + w, y + topHeight / 2);
    },

    flowChartMagneticDrum(ctx, x, y, w, h) {
        const sideWidth = w * 0.15;
        const cy = y + h / 2;
        ctx.ellipse(x + sideWidth / 2, cy, sideWidth / 2, h / 2, 0, Math.PI / 2, Math.PI * 1.5);
        ctx.lineTo(x + w - sideWidth / 2, y);
        ctx.ellipse(x + w - sideWidth / 2, cy, sideWidth / 2, h / 2, 0, -Math.PI / 2, Math.PI / 2);
        ctx.lineTo(x + sideWidth / 2, y + h);
    },

    flowChartDisplay(ctx, x, y, w, h) {
        const arcWidth = w * 0.15;
        ctx.moveTo(x, y);
        ctx.lineTo(x + w - arcWidth, y);
        ctx.arcTo(x + w + arcWidth, y + h / 2, x + w - arcWidth, y + h, h / 2);
        ctx.lineTo(x, y + h);
        ctx.closePath();
    },

    // ==================== 星形和丝带 ====================

    irregularSeal1(ctx, x, y, w, h) {
        drawRadialShape(ctx, x, y, w, h, [
            0.96, 0.46, 0.78, 0.5, 0.98, 0.42, 0.68, 0.54, 0.92,
            0.44, 0.82, 0.5, 1, 0.42, 0.7, 0.52
        ], -Math.PI / 2);
    },

    irregularSeal2(ctx, x, y, w, h) {
        drawRadialShape(ctx, x, y, w, h, [
            0.92, 0.42, 0.72, 0.5, 1, 0.38, 0.68, 0.52, 0.9, 0.44,
            0.74, 0.56, 0.96, 0.36, 0.7, 0.5, 0.88, 0.42, 0.76, 0.54
        ], -Math.PI / 2);
    },

    star4(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 4);
    },

    star5(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 5);
    },

    star6(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 6);
    },

    star7(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 7);
    },

    star8(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 8);
    },

    star10(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 10);
    },

    star12(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 12);
    },

    star16(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 16);
    },

    star24(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 24);
    },

    star32(ctx, x, y, w, h) {
        drawStar(ctx, x, y, w, h, 32);
    },

    ribbon(ctx, x, y, w, h) {
        drawRibbonTail(ctx, x, y, w, h, 'left', 0.18, 0.62, 0.32);
        drawRibbonTail(ctx, x, y, w, h, 'right', 0.18, 0.62, 0.68);
        drawRoundRect(ctx, x + w * 0.27, y + h * 0.32, w * 0.46, h * 0.5, 0.08);
    },

    ribbon2(ctx, x, y, w, h) {
        drawRoundRect(ctx, x + w * 0.27, y + h * 0.18, w * 0.46, h * 0.5, 0.08);
        drawRibbonTail(ctx, x, y, w, h, 'left', 0.38, 0.82, 0.32);
        drawRibbonTail(ctx, x, y, w, h, 'right', 0.38, 0.82, 0.68);
    },

    ellipseRibbon(ctx, x, y, w, h) {
        drawRibbonTail(ctx, x, y, w, h, 'left', 0.24, 0.62, 0.31);
        drawRibbonTail(ctx, x, y, w, h, 'right', 0.24, 0.62, 0.69);
        drawUnitPath(ctx, x, y, w, h, [
            ['M', 0.28, 0.38],
            ['C', 0.41, 0.48, 0.59, 0.48, 0.72, 0.38],
            ['L', 0.72, 0.78],
            ['C', 0.59, 0.9, 0.41, 0.9, 0.28, 0.78],
            ['Z']
        ]);
    },

    ellipseRibbon2(ctx, x, y, w, h) {
        drawRibbonTail(ctx, x, y, w, h, 'left', 0.38, 0.76, 0.31);
        drawRibbonTail(ctx, x, y, w, h, 'right', 0.38, 0.76, 0.69);
        drawUnitPath(ctx, x, y, w, h, [
            ['M', 0.28, 0.22],
            ['C', 0.41, 0.1, 0.59, 0.1, 0.72, 0.22],
            ['L', 0.72, 0.62],
            ['C', 0.59, 0.52, 0.41, 0.52, 0.28, 0.62],
            ['Z']
        ]);
    },

    verticalScroll(ctx, x, y, w, h) {
        const rollH = h * 0.1;
        const cx = x + w / 2;
        ctx.moveTo(x, y + rollH);
        ctx.arcTo(x, y, cx, y, rollH);
        ctx.arcTo(x + w, y, x + w, y + rollH, rollH);
        ctx.lineTo(x + w, y + h - rollH);
        ctx.arcTo(x + w, y + h, cx, y + h, rollH);
        ctx.arcTo(x, y + h, x, y + h - rollH, rollH);
        ctx.closePath();
    },

    horizontalScroll(ctx, x, y, w, h) {
        const rollW = w * 0.1;
        const cy = y + h / 2;
        ctx.moveTo(x + rollW, y);
        ctx.arcTo(x, y, x, cy, rollW);
        ctx.arcTo(x, y + h, x + rollW, y + h, rollW);
        ctx.lineTo(x + w - rollW, y + h);
        ctx.arcTo(x + w, y + h, x + w, cy, rollW);
        ctx.arcTo(x + w, y, x + w - rollW, y, rollW);
        ctx.closePath();
    },

    wave(ctx, x, y, w, h) {
        drawUnitPath(ctx, x, y, w, h, [
            ['M', 0, 0.28],
            ['C', 0.18, 0.06, 0.32, 0.06, 0.5, 0.28],
            ['C', 0.68, 0.5, 0.82, 0.5, 1, 0.28],
            ['L', 1, 0.76],
            ['C', 0.82, 0.98, 0.68, 0.98, 0.5, 0.76],
            ['C', 0.32, 0.54, 0.18, 0.54, 0, 0.76],
            ['Z']
        ]);
    },

    doubleWave(ctx, x, y, w, h) {
        drawUnitPath(ctx, x, y, w, h, [
            ['M', 0, 0.36],
            ['C', 0.08, 0.18, 0.17, 0.18, 0.25, 0.36],
            ['C', 0.33, 0.54, 0.42, 0.54, 0.5, 0.36],
            ['C', 0.58, 0.18, 0.67, 0.18, 0.75, 0.36],
            ['C', 0.83, 0.54, 0.92, 0.54, 1, 0.36],
            ['L', 1, 0.68],
            ['C', 0.92, 0.86, 0.83, 0.86, 0.75, 0.68],
            ['C', 0.67, 0.5, 0.58, 0.5, 0.5, 0.68],
            ['C', 0.42, 0.86, 0.33, 0.86, 0.25, 0.68],
            ['C', 0.17, 0.5, 0.08, 0.5, 0, 0.68],
            ['Z']
        ]);
    },

    // ==================== 标注 ====================

    // 矩形气泡标注（尖角在左下）
    wedgeRectCallout(ctx, x, y, w, h) {
        const bodyH = h * 0.65;  // 主体高度
        const tailW = w * 0.2;   // 尖角宽度
        const tailStartX = x + w * 0.2;

        ctx.moveTo(x, y);
        ctx.lineTo(x + w, y);
        ctx.lineTo(x + w, y + bodyH);
        ctx.lineTo(tailStartX + tailW, y + bodyH);
        ctx.lineTo(tailStartX, y + h);  // 尖角顶点
        ctx.lineTo(tailStartX, y + bodyH);
        ctx.lineTo(x, y + bodyH);
        ctx.closePath();
    },

    // 圆角矩形气泡标注
    wedgeRoundRectCallout(ctx, x, y, w, h) {
        const bodyH = h * 0.65;
        const radius = Math.min(w, h) * 0.12;
        const tailW = w * 0.2;
        const tailStartX = x + w * 0.2;

        ctx.moveTo(x + radius, y);
        ctx.lineTo(x + w - radius, y);
        ctx.arcTo(x + w, y, x + w, y + radius, radius);
        ctx.lineTo(x + w, y + bodyH - radius);
        ctx.arcTo(x + w, y + bodyH, x + w - radius, y + bodyH, radius);
        ctx.lineTo(tailStartX + tailW, y + bodyH);
        ctx.lineTo(tailStartX, y + h);  // 尖角顶点
        ctx.lineTo(tailStartX, y + bodyH);
        ctx.lineTo(x + radius, y + bodyH);
        ctx.arcTo(x, y + bodyH, x, y + bodyH - radius, radius);
        ctx.lineTo(x, y + radius);
        ctx.arcTo(x, y, x + radius, y, radius);
        ctx.closePath();
    },

    // 椭圆气泡标注
    wedgeEllipseCallout(ctx, x, y, w, h) {
        const bodyH = h * 0.6;
        const cx = x + w / 2;
        const cy = y + bodyH / 2;
        const rx = w / 2;
        const ry = bodyH / 2;

        // 椭圆主体
        ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);

        // 尖角（从椭圆底部延伸）
        const tailX = x + w * 0.35;
        const angle1 = Math.PI * 0.6;
        const angle2 = Math.PI * 0.75;
        const x1 = cx + rx * Math.cos(angle1);
        const y1 = cy + ry * Math.sin(angle1);
        const x2 = cx + rx * Math.cos(angle2);
        const y2 = cy + ry * Math.sin(angle2);

        ctx.moveTo(x1, y1);
        ctx.lineTo(tailX, y + h);
        ctx.lineTo(x2, y2);
    },

    // 云朵标注
    cloudCallout(ctx, x, y, w, h) {
        const bodyH = h * 0.7;
        const cx = x + w / 2;

        // 简化的云朵主体（用多个圆弧组成）
        const r = bodyH * 0.3;
        ctx.arc(x + w * 0.25, y + bodyH * 0.6, r, Math.PI * 0.8, Math.PI * 1.8);
        ctx.arc(cx, y + r, r * 1.2, Math.PI * 1.2, Math.PI * 1.9);
        ctx.arc(x + w * 0.75, y + bodyH * 0.5, r, Math.PI * 1.4, Math.PI * 0.3);
        ctx.arc(x + w * 0.7, y + bodyH * 0.8, r * 0.8, -Math.PI * 0.2, Math.PI * 0.5);
        ctx.arc(cx, y + bodyH * 0.85, r * 0.9, 0, Math.PI * 0.8);
        ctx.arc(x + w * 0.3, y + bodyH * 0.8, r * 0.7, Math.PI * 0.3, Math.PI);
        ctx.closePath();

        // 三个递减的小气泡
        const bubble1R = h * 0.08;
        ctx.moveTo(x + w * 0.3 + bubble1R, y + bodyH * 0.95);
        ctx.arc(x + w * 0.3, y + bodyH * 0.95, bubble1R, 0, Math.PI * 2);

        ctx.moveTo(x + w * 0.22 + bubble1R * 0.6, y + h * 0.85);
        ctx.arc(x + w * 0.22, y + h * 0.85, bubble1R * 0.6, 0, Math.PI * 2);

        ctx.moveTo(x + w * 0.18 + bubble1R * 0.35, y + h * 0.95);
        ctx.arc(x + w * 0.18, y + h * 0.95, bubble1R * 0.35, 0, Math.PI * 2);
    },

    // 线形标注1（一段线）- 主体靠右，延伸线指向左下
    callout1(ctx, x, y, w, h) {
        const boxX = x + w * 0.35;
        const boxW = w * 0.65;
        const boxH = h * 0.6;

        // 矩形主体
        ctx.rect(boxX, y, boxW, boxH);

        // 延伸线
        ctx.moveTo(boxX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 线形标注2（两段折线）
    callout2(ctx, x, y, w, h) {
        const boxX = x + w * 0.35;
        const boxW = w * 0.65;
        const boxH = h * 0.6;
        const bendX = x + w * 0.2;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(boxX, y + boxH * 0.5);
        ctx.lineTo(bendX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 线形标注3（三段折线）
    callout3(ctx, x, y, w, h) {
        const boxX = x + w * 0.35;
        const boxW = w * 0.65;
        const boxH = h * 0.6;
        const bend1X = x + w * 0.15;
        const bend1Y = y + boxH * 0.5;
        const bend2Y = y + h * 0.75;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(boxX, bend1Y);
        ctx.lineTo(bend1X, bend1Y);
        ctx.lineTo(bend1X, bend2Y);
        ctx.lineTo(x, y + h);
    },

    // 强调标注1（左侧竖线 + 一段延伸线）
    accentCallout1(ctx, x, y, w, h) {
        const accentX = x + w * 0.3;
        const boxX = x + w * 0.38;
        const boxW = w * 0.62;
        const boxH = h * 0.6;

        // 矩形主体
        ctx.rect(boxX, y, boxW, boxH);

        // 左侧强调竖线
        ctx.moveTo(accentX, y);
        ctx.lineTo(accentX, y + boxH);

        // 延伸线
        ctx.moveTo(accentX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 强调标注2
    accentCallout2(ctx, x, y, w, h) {
        const accentX = x + w * 0.3;
        const boxX = x + w * 0.38;
        const boxW = w * 0.62;
        const boxH = h * 0.6;
        const bendX = x + w * 0.15;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(accentX, y);
        ctx.lineTo(accentX, y + boxH);

        ctx.moveTo(accentX, y + boxH * 0.5);
        ctx.lineTo(bendX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 强调标注3
    accentCallout3(ctx, x, y, w, h) {
        const accentX = x + w * 0.3;
        const boxX = x + w * 0.38;
        const boxW = w * 0.62;
        const boxH = h * 0.6;
        const bend1X = x + w * 0.12;
        const bend2Y = y + h * 0.75;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(accentX, y);
        ctx.lineTo(accentX, y + boxH);

        ctx.moveTo(accentX, y + boxH * 0.5);
        ctx.lineTo(bend1X, y + boxH * 0.5);
        ctx.lineTo(bend1X, bend2Y);
        ctx.lineTo(x, y + h);
    },

    // 边框标注1（同 callout1，仅语义区分）
    borderCallout1(ctx, x, y, w, h) {
        const boxX = x + w * 0.35;
        const boxW = w * 0.65;
        const boxH = h * 0.6;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(boxX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 边框标注2
    borderCallout2(ctx, x, y, w, h) {
        const boxX = x + w * 0.35;
        const boxW = w * 0.65;
        const boxH = h * 0.6;
        const bendX = x + w * 0.2;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(boxX, y + boxH * 0.5);
        ctx.lineTo(bendX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 边框标注3
    borderCallout3(ctx, x, y, w, h) {
        const boxX = x + w * 0.35;
        const boxW = w * 0.65;
        const boxH = h * 0.6;
        const bend1X = x + w * 0.15;
        const bend2Y = y + h * 0.75;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(boxX, y + boxH * 0.5);
        ctx.lineTo(bend1X, y + boxH * 0.5);
        ctx.lineTo(bend1X, bend2Y);
        ctx.lineTo(x, y + h);
    },

    // 强调边框标注1
    accentBorderCallout1(ctx, x, y, w, h) {
        const accentX = x + w * 0.3;
        const boxX = x + w * 0.38;
        const boxW = w * 0.62;
        const boxH = h * 0.6;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(accentX, y);
        ctx.lineTo(accentX, y + boxH);

        ctx.moveTo(accentX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 强调边框标注2
    accentBorderCallout2(ctx, x, y, w, h) {
        const accentX = x + w * 0.3;
        const boxX = x + w * 0.38;
        const boxW = w * 0.62;
        const boxH = h * 0.6;
        const bendX = x + w * 0.15;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(accentX, y);
        ctx.lineTo(accentX, y + boxH);

        ctx.moveTo(accentX, y + boxH * 0.5);
        ctx.lineTo(bendX, y + boxH * 0.5);
        ctx.lineTo(x, y + h);
    },

    // 强调边框标注3
    accentBorderCallout3(ctx, x, y, w, h) {
        const accentX = x + w * 0.3;
        const boxX = x + w * 0.38;
        const boxW = w * 0.62;
        const boxH = h * 0.6;
        const bend1X = x + w * 0.12;
        const bend2Y = y + h * 0.75;

        ctx.rect(boxX, y, boxW, boxH);

        ctx.moveTo(accentX, y);
        ctx.lineTo(accentX, y + boxH);

        ctx.moveTo(accentX, y + boxH * 0.5);
        ctx.lineTo(bend1X, y + boxH * 0.5);
        ctx.lineTo(bend1X, bend2Y);
        ctx.lineTo(x, y + h);
    }
};
