/**
 * 图标集定义与Canvas矢量绘制
 * 所有图标使用Canvas绘制，无外部依赖
 */

// 图标缓存：`${iconName}-${size}` -> Canvas
const iconCache = new Map();

// 颜色定义
const COLORS = {
    green: '#63BE7B',
    yellow: '#DBAF00',  // 加深黄色
    red: '#F8696B',
    gray: '#808080',
    black: '#404040',
    pink: '#FFB6C1',
    darkGreen: '#006400',
    darkRed: '#8B0000'
};

/**
 * 获取图标Canvas（带缓存）
 * @param {string} iconName - 图标名称
 * @param {number} size - 图标尺寸
 * @returns {HTMLCanvasElement}
 */
export function getIconCanvas(iconName, size = 16) {
    const key = `${iconName}-${size}`;
    if (iconCache.has(key)) {
        return iconCache.get(key);
    }

    const canvas = document.createElement('canvas');
    canvas.width = size;
    canvas.height = size;
    const ctx = canvas.getContext('2d');

    drawIcon(ctx, iconName, size);

    iconCache.set(key, canvas);
    return canvas;
}

/**
 * 绘制图标到指定上下文
 * @param {CanvasRenderingContext2D} ctx - Canvas上下文
 * @param {string} iconName - 图标名称
 * @param {number} size - 图标尺寸
 */
export function drawIcon(ctx, iconName, size) {
    const cx = size / 2;
    const cy = size / 2;
    const r = size / 2 - 2;

    ctx.save();

    // 根据图标名称绘制
    if (iconName.startsWith('arrow')) {
        drawArrow(ctx, iconName, size);
    } else if (iconName.startsWith('light')) {
        drawTrafficLight(ctx, iconName, size);
    } else if (iconName.startsWith('flag')) {
        drawFlag(ctx, iconName, size);
    } else if (iconName.startsWith('sign')) {
        drawSign(ctx, iconName, size);
    } else if (iconName.startsWith('symbol')) {
        drawSymbol(ctx, iconName, size);
    } else if (iconName.startsWith('star')) {
        drawStar(ctx, iconName, size);
    } else if (iconName.startsWith('triangle')) {
        drawTriangle(ctx, iconName, size);
    } else if (iconName.startsWith('rating')) {
        drawRating(ctx, iconName, size);
    } else if (iconName.startsWith('circle')) {
        drawCircle(ctx, iconName, size);
    } else if (iconName.startsWith('quarter')) {
        drawQuarter(ctx, iconName, size);
    } else if (iconName.startsWith('box')) {
        drawBox(ctx, iconName, size);
    }

    ctx.restore();
}

/**
 * 绘制箭头图标
 */
function drawArrow(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const s = size * 0.35; // 箭头大小

    // 确定颜色（TiltUp/TiltDown是黄色，纯Up是绿色，纯Down是红色）
    let color, strokeColor;
    if (name.includes('Gray')) {
        color = COLORS.gray;
        strokeColor = '#555';
    } else if (name.includes('Tilt')) {
        // 斜向箭头都是黄色
        color = COLORS.yellow;
        strokeColor = '#a08000';
    } else if (name.includes('Up')) {
        color = COLORS.green;
        strokeColor = '#3d8b4f';
    } else if (name.includes('Down')) {
        color = COLORS.red;
        strokeColor = '#c44';
    } else {
        color = COLORS.yellow;
        strokeColor = '#a08000';
    }

    ctx.fillStyle = color;
    ctx.strokeStyle = strokeColor;
    ctx.lineWidth = 0.5;
    ctx.beginPath();

    if (name.includes('Up') && !name.includes('Tilt')) {
        // 向上箭头
        ctx.moveTo(cx, cy - s);
        ctx.lineTo(cx + s, cy + s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy + s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy + s);
        ctx.lineTo(cx - s * 0.3, cy + s);
        ctx.lineTo(cx - s * 0.3, cy + s * 0.5);
        ctx.lineTo(cx - s, cy + s * 0.5);
    } else if (name.includes('Down') && !name.includes('Tilt')) {
        // 向下箭头
        ctx.moveTo(cx, cy + s);
        ctx.lineTo(cx + s, cy - s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy - s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy - s);
        ctx.lineTo(cx - s * 0.3, cy - s);
        ctx.lineTo(cx - s * 0.3, cy - s * 0.5);
        ctx.lineTo(cx - s, cy - s * 0.5);
    } else if (name.includes('TiltUp')) {
        // 右上箭头（完整箭头形状，旋转45度）
        ctx.save();
        ctx.translate(cx, cy);
        ctx.rotate(Math.PI / 4);
        ctx.translate(-cx, -cy);
        ctx.moveTo(cx, cy - s);
        ctx.lineTo(cx + s, cy + s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy + s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy + s);
        ctx.lineTo(cx - s * 0.3, cy + s);
        ctx.lineTo(cx - s * 0.3, cy + s * 0.5);
        ctx.lineTo(cx - s, cy + s * 0.5);
        ctx.restore();
    } else if (name.includes('TiltDown')) {
        // 右下箭头（完整箭头形状，旋转-45度）
        ctx.save();
        ctx.translate(cx, cy);
        ctx.rotate(-Math.PI / 4);
        ctx.translate(-cx, -cy);
        ctx.moveTo(cx, cy + s);
        ctx.lineTo(cx + s, cy - s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy - s * 0.5);
        ctx.lineTo(cx + s * 0.3, cy - s);
        ctx.lineTo(cx - s * 0.3, cy - s);
        ctx.lineTo(cx - s * 0.3, cy - s * 0.5);
        ctx.lineTo(cx - s, cy - s * 0.5);
        ctx.restore();
    } else if (name.includes('Side')) {
        // 水平箭头
        ctx.moveTo(cx + s, cy);
        ctx.lineTo(cx - s * 0.5, cy - s);
        ctx.lineTo(cx - s * 0.5, cy - s * 0.3);
        ctx.lineTo(cx - s, cy - s * 0.3);
        ctx.lineTo(cx - s, cy + s * 0.3);
        ctx.lineTo(cx - s * 0.5, cy + s * 0.3);
        ctx.lineTo(cx - s * 0.5, cy + s);
    }

    ctx.closePath();
    ctx.fill();
    ctx.stroke();
}

/**
 * 绘制交通灯图标
 */
function drawTrafficLight(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const r = size * 0.35;

    let color;
    if (name.includes('Green')) {
        color = COLORS.green;
    } else if (name.includes('Yellow')) {
        color = COLORS.yellow;
    } else if (name.includes('Red')) {
        color = COLORS.red;
    } else {
        color = COLORS.black;
    }

    if (name.includes('2')) {
        // 带边框版本：圆角正方形外框 + 内部圆形
        const boxSize = size * 0.8;
        const boxX = (size - boxSize) / 2;
        const boxY = (size - boxSize) / 2;
        const radius = size * 0.15; // 圆角半径

        // 绘制圆角正方形外框
        ctx.fillStyle = '#404040';
        ctx.beginPath();
        ctx.roundRect(boxX, boxY, boxSize, boxSize, radius);
        ctx.fill();

        // 绘制内部彩色圆形
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.arc(cx, cy, r * 0.85, 0, Math.PI * 2);
        ctx.fill();
    } else {
        // 纯色版本
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.arc(cx, cy, r, 0, Math.PI * 2);
        ctx.fill();
    }
}

/**
 * 绘制旗帜图标
 */
function drawFlag(ctx, name, size) {
    const x = size * 0.25;
    const y = size * 0.15;
    const w = size * 0.5;
    const h = size * 0.4;

    let color;
    if (name.includes('Green')) {
        color = COLORS.green;
    } else if (name.includes('Yellow')) {
        color = COLORS.yellow;
    } else {
        color = COLORS.red;
    }

    // 旗杆
    ctx.fillStyle = '#666';
    ctx.fillRect(x - 1, y, 2, size * 0.7);

    // 旗面
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + w, y + h * 0.5);
    ctx.lineTo(x, y + h);
    ctx.closePath();
    ctx.fill();
}

/**
 * 绘制标志图标
 */
function drawSign(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const r = size * 0.35;

    let color;
    if (name.includes('Green')) {
        color = COLORS.green;
    } else if (name.includes('Yellow')) {
        color = COLORS.yellow;
    } else {
        color = COLORS.red;
    }

    ctx.fillStyle = color;
    ctx.beginPath();

    if (name.includes('Red')) {
        // 红色：菱形（钻石形）
        const r2 = size * 0.4;
        ctx.moveTo(cx, cy - r2);      // 上
        ctx.lineTo(cx + r2, cy);      // 右
        ctx.lineTo(cx, cy + r2);      // 下
        ctx.lineTo(cx - r2, cy);      // 左
    } else if (name.includes('Yellow')) {
        // 黄色：三角形（警告标志）
        ctx.moveTo(cx, cy - r);
        ctx.lineTo(cx + r, cy + r * 0.7);
        ctx.lineTo(cx - r, cy + r * 0.7);
    } else {
        // 绿色：圆形（通行标志）
        ctx.arc(cx, cy, r, 0, Math.PI * 2);
    }

    ctx.closePath();
    ctx.fill();
}

/**
 * 绘制符号图标
 */
function drawSymbol(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const noCircle = name.includes('2'); // 2版本是无圈
    const r = noCircle ? size * 0.45 : size * 0.35 + 2; // 无圈版本更大

    if (name.includes('Check')) {
        if (!noCircle) {
            // 有圈：绿色圆形背景 + 白色对勾
            ctx.fillStyle = COLORS.green;
            ctx.beginPath();
            ctx.arc(cx, cy, r, 0, Math.PI * 2);
            ctx.fill();
            ctx.strokeStyle = '#fff';
            ctx.lineWidth = 1.5;
        } else {
            // 无圈：绿色对勾
            ctx.strokeStyle = COLORS.green;
            ctx.lineWidth = 2;
        }
        ctx.beginPath();
        ctx.moveTo(cx - r * 0.5, cy);
        ctx.lineTo(cx - r * 0.15, cy + r * 0.45);
        ctx.lineTo(cx + r * 0.5, cy - r * 0.35);
        ctx.stroke();
    } else if (name.includes('Exclaim')) {
        if (!noCircle) {
            // 有圈：黄色圆形背景 + 白色感叹号
            ctx.fillStyle = COLORS.yellow;
            ctx.beginPath();
            ctx.arc(cx, cy, r, 0, Math.PI * 2);
            ctx.fill();
            // 白色竖线
            ctx.fillStyle = '#fff';
            ctx.beginPath();
            ctx.roundRect(cx - 1.5, cy - r * 0.6, 3, r * 0.55, 1);
            ctx.fill();
            // 白色圆点
            ctx.beginPath();
            ctx.arc(cx, cy + r * 0.45, 2, 0, Math.PI * 2);
            ctx.fill();
        } else {
            // 无圈：黄色感叹号（线条）
            ctx.strokeStyle = COLORS.yellow;
            ctx.lineCap = 'round';
            ctx.lineWidth = 2.5;
            ctx.beginPath();
            ctx.moveTo(cx, cy - r * 0.6);
            ctx.lineTo(cx, cy + r * 0.1);
            ctx.stroke();
            // 圆点
            ctx.fillStyle = COLORS.yellow;
            ctx.beginPath();
            ctx.arc(cx, cy + r * 0.5, 2, 0, Math.PI * 2);
            ctx.fill();
        }
    } else if (name.includes('X')) {
        if (!noCircle) {
            // 有圈：红色圆形背景 + 白色X
            ctx.fillStyle = COLORS.red;
            ctx.beginPath();
            ctx.arc(cx, cy, r, 0, Math.PI * 2);
            ctx.fill();
            ctx.strokeStyle = '#fff';
            ctx.lineWidth = 1.5;
        } else {
            // 无圈：红色X
            ctx.strokeStyle = COLORS.red;
            ctx.lineWidth = 2;
        }
        ctx.beginPath();
        ctx.moveTo(cx - r * 0.4, cy - r * 0.4);
        ctx.lineTo(cx + r * 0.4, cy + r * 0.4);
        ctx.moveTo(cx + r * 0.4, cy - r * 0.4);
        ctx.lineTo(cx - r * 0.4, cy + r * 0.4);
        ctx.stroke();
    }
}

/**
 * 绘制星星图标
 */
function drawStar(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const outerR = size * 0.4;
    const innerR = size * 0.2;

    ctx.fillStyle = '#FFD700'; // 金色
    ctx.strokeStyle = '#B8860B';
    ctx.lineWidth = 1;

    ctx.beginPath();
    for (let i = 0; i < 5; i++) {
        const outerAngle = (i * 2 * Math.PI / 5) - Math.PI / 2;
        const innerAngle = outerAngle + Math.PI / 5;

        const ox = cx + outerR * Math.cos(outerAngle);
        const oy = cy + outerR * Math.sin(outerAngle);
        const ix = cx + innerR * Math.cos(innerAngle);
        const iy = cy + innerR * Math.sin(innerAngle);

        if (i === 0) ctx.moveTo(ox, oy);
        else ctx.lineTo(ox, oy);
        ctx.lineTo(ix, iy);
    }
    ctx.closePath();

    if (name.includes('Full')) {
        ctx.fill();
        ctx.stroke();
    } else if (name.includes('Half')) {
        ctx.stroke();
        ctx.save();
        ctx.clip();
        ctx.fillRect(0, 0, cx, size);
        ctx.restore();
    } else {
        ctx.stroke();
    }
}

/**
 * 绘制三角形图标
 */
function drawTriangle(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const s = size * 0.35;

    if (name.includes('Up')) {
        ctx.fillStyle = COLORS.green;
        ctx.beginPath();
        ctx.moveTo(cx, cy - s);
        ctx.lineTo(cx + s, cy + s * 0.7);
        ctx.lineTo(cx - s, cy + s * 0.7);
        ctx.closePath();
        ctx.fill();
    } else if (name.includes('Down')) {
        ctx.fillStyle = COLORS.red;
        ctx.beginPath();
        ctx.moveTo(cx, cy + s);
        ctx.lineTo(cx + s, cy - s * 0.7);
        ctx.lineTo(cx - s, cy - s * 0.7);
        ctx.closePath();
        ctx.fill();
    } else {
        // 横线
        ctx.fillStyle = COLORS.yellow;
        ctx.fillRect(cx - s, cy - 2, s * 2, 4);
    }
}

/**
 * 绘制评分图标
 */
function drawRating(ctx, name, size) {
    const bars = parseInt(name.replace('rating', '')) || 0;
    // 根据评级数判断总柱子数（rating5存在则是5Rating，否则是4Rating）
    const maxBars = bars <= 4 ? 4 : 5;
    const barWidth = size * (maxBars === 5 ? 0.12 : 0.15);
    const barGap = size * 0.04;
    const startX = (size - (maxBars * barWidth + (maxBars - 1) * barGap)) / 2;

    for (let i = 0; i < maxBars; i++) {
        const x = startX + i * (barWidth + barGap);
        const height = size * (0.25 + i * 0.12);
        const y = size - height - size * 0.1;

        if (i < bars) {
            ctx.fillStyle = '#4472C4'; // 蓝色
        } else {
            ctx.fillStyle = '#ddd';
        }
        ctx.fillRect(x, y, barWidth, height);
    }
}

/**
 * 绘制圆形图标（红到黑系列）
 */
function drawCircle(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const r = size * 0.35;

    let color;
    if (name.includes('Red')) {
        color = COLORS.red;
    } else if (name.includes('Pink')) {
        color = COLORS.pink;
    } else if (name.includes('Gray')) {
        color = COLORS.gray;
    } else {
        color = COLORS.black;
    }

    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.fill();
}

/**
 * 绘制四分图标
 */
function drawQuarter(ctx, name, size) {
    const cx = size / 2;
    const cy = size / 2;
    const r = size * 0.35;
    const quarters = parseInt(name.replace('quarter', '')) || 0;

    // 背景圆
    ctx.fillStyle = '#ddd';
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.fill();

    // 填充部分（黑色）
    if (quarters > 0) {
        ctx.fillStyle = '#333';
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + (quarters / 4) * Math.PI * 2);
        ctx.closePath();
        ctx.fill();
    }

    // 边框
    ctx.strokeStyle = '#666';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.stroke();
}

/**
 * 绘制方块图标（2x2四宫格）
 */
function drawBox(ctx, name, size) {
    const boxes = parseInt(name.replace('box', '')) || 0;
    const gridSize = size * 0.32; // 每个格子大小
    const gap = size * 0.06;      // 格子间距
    const totalSize = gridSize * 2 + gap;
    const startX = (size - totalSize) / 2;
    const startY = (size - totalSize) / 2;

    // 2x2 网格，填充顺序：左上、右上、左下、右下
    const positions = [
        [0, 0], [1, 0], [0, 1], [1, 1]
    ];

    for (let i = 0; i < 4; i++) {
        const [col, row] = positions[i];
        const x = startX + col * (gridSize + gap);
        const y = startY + row * (gridSize + gap);

        if (i < boxes) {
            ctx.fillStyle = '#4472C4'; // 蓝色
        } else {
            ctx.fillStyle = '#ddd';
        }
        ctx.fillRect(x, y, gridSize, gridSize);
        ctx.strokeStyle = '#999';
        ctx.lineWidth = 0.5;
        ctx.strokeRect(x, y, gridSize, gridSize);
    }
}

/**
 * 清除图标缓存
 */
export function clearIconCache() {
    iconCache.clear();
}
