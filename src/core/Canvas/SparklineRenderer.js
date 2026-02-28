/**
 * 迷你图渲染模块
 * 高性能设计：只渲染视图内的迷你图，使用缓存避免重复计算
 */

const SPARKLINE_COLOR_FALLBACKS = {
    series: '#376092',
    negative: '#D00000',
    axis: '#000000',
    markers: '#D00000',
    first: '#D00000',
    last: '#D00000',
    high: '#D00000',
    low: '#D00000'
};

const getColor = (group, key) => group?.colors?.[key] ?? SPARKLINE_COLOR_FALLBACKS[key];

/**
 * 渲染所有视图内的迷你图
 * @param {Object} sheet - 工作表对象
 */
export function rSparklines(sheet) {
    if (!sheet.Sparkline || sheet.Sparkline.map.size === 0) return;

    const { rowsIndex, colsIndex } = sheet.vi;

    // 遍历视图内的所有单元格，累加位置
    let y = sheet.headHeight;
    for (let ri = 0; ri < rowsIndex.length; ri++) {
        const r = rowsIndex[ri];
        const rowHeight = sheet.getRow(r).height;

        let x = sheet.indexWidth;
        for (let ci = 0; ci < colsIndex.length; ci++) {
            const c = colsIndex[ci];
            const colWidth = sheet.getCol(c).width;

            // 检查该位置是否有迷你图
            const sparkline = sheet.Sparkline.get({ r, c });
            if (sparkline) {
                // 获取数据
                const data = sparkline.getData();
                if (data && data.length > 0) {
                    // 过滤 null 值并准备绘制
                    const validData = data.filter(v => v !== null);
                    if (validData.length > 0) {
                        // 绘制迷你图
                        _rSparkline.call(this, x, y, colWidth, rowHeight, sparkline, validData, sheet);
                    }
                }
            }

            x += colWidth;
        }

        y += rowHeight;
    }
}

/**
 * 绘制单个迷你图
 * @private
 */
function _rSparkline(x, y, w, h, sparkline, data, sheet) {
    const { group } = sparkline;
    const ctx = this.bc;

    // 计算内边距
    let paddingTop, paddingBottom, paddingLeft, paddingRight;

    if (group.type === 'line') {
        // 折线图左右留15%空间，上下3px
        paddingLeft = w * 0.15;
        paddingRight = w * 0.15;
        paddingTop = 3;
        paddingBottom = 3;
    } else if (group.type === 'column') {
        // 柱状图上下固定2px，左右3px
        paddingTop = 2;
        paddingBottom = 2;
        paddingLeft = 3;
        paddingRight = 3;
    } else {
        // 盈亏图上下各留少量空间
        paddingTop = 2;
        paddingBottom = 2;
        paddingLeft = 2;
        paddingRight = 2;
    }

    const chartX = x + paddingLeft;
    const chartY = y + paddingTop;
    const chartW = w - paddingLeft - paddingRight;
    const chartH = h - paddingTop - paddingBottom;

    if (chartW <= 0 || chartH <= 0) return;

    // ???????????
    const dataMax = Math.max(...data);
    const dataMin = Math.min(...data);
    const maxVal = group.manualMax ?? dataMax;
    const minVal = group.manualMin ?? dataMin;
    const range = maxVal - minVal || 1; // ???? 0

    ctx.save();

    // ??????
    switch (group.type) {
        case 'line':
            _rSparklineLine.call(this, ctx, chartX, chartY, chartW, chartH, data, minVal, range, group);
            break;
        case 'column':
            _rSparklineColumn.call(this, ctx, chartX, chartY, chartW, chartH, data, minVal, maxVal, range, group, dataMin, dataMax);
            break;
        case 'stacked': // Excel中盈亏图类型为stacked
            _rSparklineWinLoss.call(this, ctx, chartX, chartY, chartW, chartH, data, group);
            break;
    }

    ctx.restore();
}

function _rSparklineLine(ctx, x, y, w, h, data, minVal, range, group) {
    const count = data.length;
    const stepX = w / (count - 1 || 1);

    ctx.beginPath();
    ctx.strokeStyle = getColor(group, 'series');
    const lineWeight = group.lineWeight || 1;
    ctx.lineWidth = this.px(lineWeight);

    // ????
    data.forEach((val, i) => {
        const px = x + i * stepX;
        const py = y + h - ((val - minVal) / range) * h;

        if (i === 0) {
            ctx.moveTo(px, py);
        } else {
            ctx.lineTo(px, py);
        }
    });

    ctx.stroke();

    // ?????
    if (group.markers || group.first || group.last || group.high || group.low) {
        const dataMax = Math.max(...data);
        const dataMin = Math.min(...data);
        const maxIdx = data.indexOf(dataMax);
        const minIdx = data.indexOf(dataMin);

        const markerColor = getColor(group, 'markers');
        const firstColor = getColor(group, 'first');
        const lastColor = getColor(group, 'last');
        const highColor = getColor(group, 'high');
        const lowColor = getColor(group, 'low');

        data.forEach((val, i) => {
            const px = x + i * stepX;
            const py = y + h - ((val - minVal) / range) * h;

            let shouldDraw = false;
            let color = markerColor;

            if (group.markers) {
                shouldDraw = true;
            }
            if (group.first && i === 0) {
                shouldDraw = true;
                color = firstColor;
            }
            if (group.last && i === count - 1) {
                shouldDraw = true;
                color = lastColor;
            }
            if (group.high && i === maxIdx) {
                shouldDraw = true;
                color = highColor;
            }
            if (group.low && i === minIdx) {
                shouldDraw = true;
                color = lowColor;
            }

            if (shouldDraw) {
                ctx.beginPath();
                ctx.fillStyle = color;
                ctx.arc(px, py, 1.5, 0, Math.PI * 2);
                ctx.fill();
            }
        });
    }
}

function _rSparklineColumn(ctx, x, y, w, h, data, minVal, maxVal, range, group, dataMin, dataMax) {
    const count = data.length;
    const colWidth = w / count;
    const gap = Math.max(0.5, colWidth * 0.1);
    const barWidth = Math.max(1, colWidth - gap);

    // ??????
    const zeroY = maxVal <= 0 ? y : minVal >= 0 ? y + h : y + h - ((-minVal) / range) * h;

    const axisColor = getColor(group, 'axis');
    const seriesColor = getColor(group, 'series');
    const negativeColor = getColor(group, 'negative');
    const firstColor = getColor(group, 'first');
    const lastColor = getColor(group, 'last');
    const highColor = getColor(group, 'high');
    const lowColor = getColor(group, 'low');

    // ????
    if (minVal < 0 && maxVal > 0) {
        ctx.beginPath();
        ctx.strokeStyle = axisColor;
        ctx.lineWidth = this.px(1);
        ctx.moveTo(x, zeroY);
        ctx.lineTo(x + w, zeroY);
        ctx.stroke();
    }

    // 绘制柱子
    data.forEach((val, i) => {
        if (val === 0 && minVal === 0 && maxVal === 0) return; // 全是0跳过

        const px = x + i * colWidth + gap / 2;

        // 计算柱子高度和位置
        let barHeight, py;
        if (minVal >= 0) {
            // 全正数：从minVal作为基准，最小高度1px
            barHeight = Math.max(1, ((val - minVal) / range) * h);
            py = zeroY - barHeight;
        } else if (maxVal <= 0) {
            // 全负数：从maxVal作为基准，最小高度1px
            barHeight = Math.max(1, (Math.abs(val - maxVal) / range) * h);
            py = zeroY;
        } else {
            // 有正有负：从0作为基准，最小高度1px
            barHeight = Math.max(1, (Math.abs(val) / range) * h);
            py = val >= 0 ? zeroY - barHeight : zeroY;
        }

        // 颜色优先级：series < negative < first/last < high/low
        let color = seriesColor;
        if (val < 0 && group.negative) color = negativeColor;
        if (group.first && i === 0) color = firstColor;
        if (group.last && i === count - 1) color = lastColor;
        if (group.high && val === dataMax) color = highColor;
        if (group.low && val === dataMin) color = lowColor;

        ctx.fillStyle = color;
        ctx.fillRect(px, py, barWidth, barHeight);
    });
}

function _rSparklineWinLoss(ctx, x, y, w, h, data, group) {
    const count = data.length;
    const colWidth = w / count;
    const gap = Math.max(0.5, colWidth * 0.1);
    const barWidth = Math.max(1, colWidth - gap);
    const zeroY = y + h / 2;
    const barHeight = h / 2 - 1;

    const axisColor = getColor(group, 'axis');
    const positiveColor = getColor(group, 'series');
    const negativeColor = getColor(group, 'negative');
    const firstColor = getColor(group, 'first');
    const lastColor = getColor(group, 'last');
    const highColor = getColor(group, 'high');
    const lowColor = getColor(group, 'low');

    // 查找最大最小值索引（用于标记）
    const dataMax = Math.max(...data);
    const dataMin = Math.min(...data);

    // 绘制轴线
    ctx.beginPath();
    ctx.strokeStyle = axisColor;
    ctx.lineWidth = this.px(1);
    ctx.moveTo(x, zeroY);
    ctx.lineTo(x + w, zeroY);
    ctx.stroke();

    // 绘制柱子
    data.forEach((val, i) => {
        if (val === 0) return;

        const px = x + i * colWidth + gap / 2;
        const py = val > 0 ? zeroY - barHeight : zeroY;

        // 颜色优先级：positive/negative < first/last < high/low
        let color = val > 0 ? positiveColor : negativeColor;
        if (group.first && i === 0) color = firstColor;
        if (group.last && i === count - 1) color = lastColor;
        if (group.high && val === dataMax) color = highColor;
        if (group.low && val === dataMin) color = lowColor;

        ctx.fillStyle = color;
        ctx.fillRect(px, py, barWidth, barHeight);
    });
}
