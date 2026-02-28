const COLORS = {
    blue: '#4472C4',
    lightBlue: '#5B9BD5',
    orange: '#ED7D31',
    gray: '#A5A5A5',
    yellow: '#FFC000',
    green: '#70AD47',
    purple: '#7030A0'
};

const CHART = {
    left: 6,
    right: 114,
    top: 4,
    bottom: 76
};

const WIDTH = CHART.right - CHART.left;
const HEIGHT = CHART.bottom - CHART.top;

const fmt = (value) => {
    const rounded = Math.round(value * 10) / 10;
    return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1);
};

const x = (p) => fmt(CHART.left + p * WIDTH);
const y = (p) => fmt(CHART.bottom - p * HEIGHT);

const AXIS = `<line x1="${CHART.left}" y1="${CHART.bottom}" x2="${CHART.right}" y2="${CHART.bottom}" stroke="#c6cbd3" stroke-width="1"/><line x1="${CHART.left}" y1="${CHART.top}" x2="${CHART.left}" y2="${CHART.bottom}" stroke="#c6cbd3" stroke-width="1"/>`;

const toPoints = (pts) => pts.map(([px, py]) => `${fmt(px)},${fmt(py)}`).join(' ');
const polyline = (pts, color, width = 2) => `<polyline points="${toPoints(pts)}" fill="none" stroke="${color}" stroke-width="${width}" stroke-linecap="round" stroke-linejoin="round"/>`;
const polygon = (pts, fill, opacity) => `<polygon points="${toPoints(pts)}" fill="${fill}" opacity="${opacity}"/>`;
const markers = (pts, color, r = 2.4) => pts.map(([px, py]) => `<circle cx="${fmt(px)}" cy="${fmt(py)}" r="${r}" fill="${color}"/>`).join('');
const smoothPath = (pts) => {
    if (pts.length < 2) return '';
    let d = `M ${fmt(pts[0][0])} ${fmt(pts[0][1])}`;
    for (let i = 0; i < pts.length - 1; i++) {
        const p0 = pts[i - 1] || pts[i];
        const p1 = pts[i];
        const p2 = pts[i + 1];
        const p3 = pts[i + 2] || p2;
        const c1x = p1[0] + (p2[0] - p0[0]) / 6;
        const c1y = p1[1] + (p2[1] - p0[1]) / 6;
        const c2x = p2[0] - (p3[0] - p1[0]) / 6;
        const c2y = p2[1] - (p3[1] - p1[1]) / 6;
        d += ` C ${fmt(c1x)} ${fmt(c1y)} ${fmt(c2x)} ${fmt(c2y)} ${fmt(p2[0])} ${fmt(p2[1])}`;
    }
    return d;
};

const wrapSvg = (inner) => `<svg viewBox="0 0 120 80" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">${inner}</svg>`;

const LINE_VALUES = [0.22, 0.65, 0.45, 0.88, 0.6];
const LINE_VALUES_ALT = [0.18, 0.42, 0.3, 0.5, 0.36];
const LINE_VALUES_TOP = [0.55, 0.72, 0.62, 0.82, 0.74];

const linePoints = (values) => values.map((v, i) => [CHART.left + (WIDTH * i) / (values.length - 1), CHART.bottom - v * HEIGHT]);
const areaFill = (points, color, opacity) => polygon([...points, [CHART.right, CHART.bottom], [CHART.left, CHART.bottom]], color, opacity);
const areaBetween = (lowerPoints, upperPoints, color, opacity) => polygon([...upperPoints, ...lowerPoints.slice().reverse()], color, opacity);

const bar = (centerPct, widthPct, value, color) => {
    const barW = WIDTH * widthPct;
    const centerX = CHART.left + centerPct * WIDTH;
    const barX = centerX - barW / 2;
    const barY = CHART.bottom - value * HEIGHT;
    const barH = CHART.bottom - barY;
    return `<rect x="${fmt(barX)}" y="${fmt(barY)}" width="${fmt(barW)}" height="${fmt(barH)}" fill="${color}"/>`;
};

const stackedBar = (centerPct, widthPct, values, colors) => {
    const barW = WIDTH * widthPct;
    const centerX = CHART.left + centerPct * WIDTH;
    const barX = centerX - barW / 2;
    let baseY = CHART.bottom;
    return values.map((value, idx) => {
        const segH = value * HEIGHT;
        const segY = baseY - segH;
        baseY = segY;
        return `<rect x="${fmt(barX)}" y="${fmt(segY)}" width="${fmt(barW)}" height="${fmt(segH)}" fill="${colors[idx]}"/>`;
    }).join('');
};

const barH = (centerPct, heightPct, value, color) => {
    const barHgt = HEIGHT * heightPct;
    const centerY = CHART.top + centerPct * HEIGHT;
    const barY = centerY - barHgt / 2;
    const barW = value * WIDTH;
    return `<rect x="${CHART.left}" y="${fmt(barY)}" width="${fmt(barW)}" height="${fmt(barHgt)}" fill="${color}"/>`;
};

const stackedBarH = (centerPct, heightPct, values, colors) => {
    const barHgt = HEIGHT * heightPct;
    const centerY = CHART.top + centerPct * HEIGHT;
    const barY = centerY - barHgt / 2;
    let baseX = CHART.left;
    return values.map((value, idx) => {
        const segW = value * WIDTH;
        const seg = `<rect x="${fmt(baseX)}" y="${fmt(barY)}" width="${fmt(segW)}" height="${fmt(barHgt)}" fill="${colors[idx]}"/>`;
        baseX += segW;
        return seg;
    }).join('');
};

const floatingBar = (centerPct, widthPct, fromVal, toVal, color) => {
    const barW = WIDTH * widthPct;
    const centerX = CHART.left + centerPct * WIDTH;
    const barX = centerX - barW / 2;
    const topVal = Math.max(fromVal, toVal);
    const bottomVal = Math.min(fromVal, toVal);
    const barY = CHART.bottom - topVal * HEIGHT;
    const barH = (topVal - bottomVal) * HEIGHT;
    return `<rect x="${fmt(barX)}" y="${fmt(barY)}" width="${fmt(barW)}" height="${fmt(barH)}" fill="${color}"/>`;
};

const ringSegments = (r, width, segments, colors) => {
    const circumference = 2 * Math.PI * r;
    let offset = 0;
    return segments.map((ratio, idx) => {
        const len = ratio * circumference;
        const dash = `${len.toFixed(2)} ${(circumference - len).toFixed(2)}`;
        const circle = `<circle cx="60" cy="40" r="${r}" fill="none" stroke="${colors[idx]}" stroke-width="${width}" stroke-dasharray="${dash}" stroke-dashoffset="${(-offset).toFixed(2)}" transform="rotate(-90 60 40)"/>`;
        offset += len;
        return circle;
    }).join('');
};

const radarPoints = (radius) => {
    const angles = [-90, -18, 54, 126, 198].map((deg) => (deg * Math.PI) / 180);
    return angles.map((rad) => [60 + Math.cos(rad) * radius, 40 + Math.sin(rad) * radius]);
};

const radarGrid = () => {
    const rings = [38, 28, 18];
    const ringLines = rings.map((r) => `<polygon points="${toPoints(radarPoints(r))}" fill="none" stroke="#d7dbe2" stroke-width="1"/>`).join('');
    const axes = radarPoints(38).map(([px, py]) => `<line x1="60" y1="40" x2="${fmt(px)}" y2="${fmt(py)}" stroke="#d7dbe2" stroke-width="1"/>`).join('');
    return `${ringLines}${axes}`;
};

const radarShape = (radius, ratios) => radarPoints(radius).map(([px, py], idx) => {
    const ratio = ratios[idx] ?? 1;
    return [60 + (px - 60) * ratio, 40 + (py - 40) * ratio];
});

function barClustered() {
    return wrapSvg(`${AXIS}
        ${bar(0.1, 0.08, 0.45, COLORS.blue)}
        ${bar(0.3, 0.08, 0.7, COLORS.orange)}
        ${bar(0.5, 0.08, 0.35, COLORS.gray)}
        ${bar(0.7, 0.08, 0.82, COLORS.green)}
        ${bar(0.9, 0.08, 0.6, COLORS.yellow)}
    `);
}

function barStacked() {
    return wrapSvg(`${AXIS}
        ${stackedBar(0.2, 0.1, [0.35, 0.2], [COLORS.blue, COLORS.orange])}
        ${stackedBar(0.5, 0.1, [0.28, 0.3], [COLORS.gray, COLORS.yellow])}
        ${stackedBar(0.8, 0.1, [0.22, 0.32], [COLORS.green, COLORS.lightBlue])}
    `);
}

function barPercent() {
    return wrapSvg(`${AXIS}
        ${stackedBar(0.2, 0.1, [0.5, 0.3, 0.2], [COLORS.blue, COLORS.orange, COLORS.gray])}
        ${stackedBar(0.5, 0.1, [0.35, 0.4, 0.25], [COLORS.green, COLORS.yellow, COLORS.lightBlue])}
        ${stackedBar(0.8, 0.1, [0.4, 0.35, 0.25], [COLORS.orange, COLORS.blue, COLORS.gray])}
    `);
}

function histogramBasic() {
    return wrapSvg(`${AXIS}
        ${bar(0.12, 0.18, 0.4, COLORS.blue)}
        ${bar(0.3, 0.18, 0.65, COLORS.orange)}
        ${bar(0.48, 0.18, 0.32, COLORS.gray)}
        ${bar(0.66, 0.18, 0.82, COLORS.green)}
        ${bar(0.84, 0.18, 0.55, COLORS.yellow)}
    `);
}

function barHClustered() {
    return wrapSvg(`${AXIS}
        ${barH(0.2, 0.12, 0.7, COLORS.blue)}
        ${barH(0.5, 0.12, 0.85, COLORS.orange)}
        ${barH(0.8, 0.12, 0.55, COLORS.gray)}
    `);
}

function barHStacked() {
    return wrapSvg(`${AXIS}
        ${stackedBarH(0.25, 0.12, [0.35, 0.25, 0.18], [COLORS.blue, COLORS.orange, COLORS.gray])}
        ${stackedBarH(0.55, 0.12, [0.25, 0.3, 0.25], [COLORS.green, COLORS.yellow, COLORS.lightBlue])}
        ${stackedBarH(0.85, 0.12, [0.3, 0.32, 0.2], [COLORS.orange, COLORS.blue, COLORS.gray])}
    `);
}

function barHPercent() {
    return wrapSvg(`${AXIS}
        ${stackedBarH(0.25, 0.12, [0.4, 0.35, 0.25], [COLORS.blue, COLORS.orange, COLORS.gray])}
        ${stackedBarH(0.55, 0.12, [0.3, 0.45, 0.25], [COLORS.green, COLORS.yellow, COLORS.lightBlue])}
        ${stackedBarH(0.85, 0.12, [0.35, 0.4, 0.25], [COLORS.orange, COLORS.blue, COLORS.gray])}
    `);
}

function lineBasic(withMarkers) {
    const primary = linePoints(LINE_VALUES);
    const secondary = linePoints(LINE_VALUES_ALT);
    return wrapSvg(`${AXIS}
        ${polyline(secondary, COLORS.orange, 2.0)}
        ${polyline(primary, COLORS.blue, 2.1)}
        ${withMarkers ? `${markers(secondary, COLORS.orange, 2.4)}${markers(primary, COLORS.blue, 2.6)}` : ''}
    `);
}

function lineStacked(withMarkers) {
    const baseValues = LINE_VALUES_ALT;
    const topValues = LINE_VALUES.map((v, i) => Math.min(0.95, v + baseValues[i]));
    const base = linePoints(baseValues);
    const top = linePoints(topValues);
    return wrapSvg(`${AXIS}
        ${polyline(base, COLORS.blue, 2.0)}
        ${polyline(top, COLORS.orange, 2.0)}
        ${withMarkers ? `${markers(base, COLORS.blue, 2.4)}${markers(top, COLORS.orange, 2.4)}` : ''}
    `);
}

function linePercent(withMarkers) {
    const baseRaw = LINE_VALUES_ALT.map((v) => Math.max(0.18, Math.min(0.72, v)));
    const topRaw = LINE_VALUES_TOP.map((v) => Math.max(0.32, Math.min(0.88, v)));
    const basePct = baseRaw.map((v, i) => {
        const sum = v + topRaw[i];
        const ratio = v / sum;
        const amplified = 0.5 + (ratio - 0.5) * 1.5;
        return Math.max(0.18, Math.min(0.82, amplified));
    });
    const base = linePoints(basePct);
    const top = linePoints(basePct.map(() => 0.98));
    return wrapSvg(`${AXIS}
        ${polyline(base, COLORS.blue, 2.0)}
        ${polyline(top, COLORS.orange, 2.0)}
        ${withMarkers ? `${markers(base, COLORS.blue, 2.4)}${markers(top, COLORS.orange, 2.4)}` : ''}
    `);
}

function areaBasic() {
    const points = linePoints(LINE_VALUES);
    return wrapSvg(`${AXIS}${areaFill(points, COLORS.blue, 0.25)}${polyline(points, COLORS.blue, 2.1)}`);
}

function areaStacked() {
    const baseValues = LINE_VALUES_ALT;
    const topValues = LINE_VALUES.map((v, i) => Math.min(0.95, v + baseValues[i]));
    const base = linePoints(baseValues);
    const top = linePoints(topValues);
    return wrapSvg(`${AXIS}
        ${areaFill(base, COLORS.lightBlue, 0.25)}
        ${areaBetween(base, top, COLORS.orange, 0.25)}
        ${polyline(base, COLORS.lightBlue, 2.0)}
        ${polyline(top, COLORS.orange, 2.0)}
    `);
}

function areaPercent() {
    const baseRaw = LINE_VALUES_ALT;
    const topRaw = LINE_VALUES;
    const basePct = baseRaw.map((v, i) => {
        const sum = v + topRaw[i];
        return Math.max(0.12, Math.min(0.88, v / sum));
    });
    const base = linePoints(basePct);
    const top = linePoints(basePct.map(() => 0.98));
    return wrapSvg(`${AXIS}
        ${areaFill(base, COLORS.blue, 0.25)}
        ${areaBetween(base, top, COLORS.orange, 0.25)}
        ${polyline(base, COLORS.blue, 2.0)}
        ${polyline(top, COLORS.orange, 2.0)}
    `);
}

function scatterBasic() {
    const points = [
        [CHART.left + WIDTH * 0.08, CHART.bottom - HEIGHT * 0.25],
        [CHART.left + WIDTH * 0.28, CHART.bottom - HEIGHT * 0.6],
        [CHART.left + WIDTH * 0.45, CHART.bottom - HEIGHT * 0.4],
        [CHART.left + WIDTH * 0.6, CHART.bottom - HEIGHT * 0.8],
        [CHART.left + WIDTH * 0.78, CHART.bottom - HEIGHT * 0.55],
        [CHART.left + WIDTH * 0.92, CHART.bottom - HEIGHT * 0.7]
    ];
    return wrapSvg(`${AXIS}${markers(points, COLORS.blue, 2.8)}`);
}

function scatterSmooth(withMarkers) {
    const points = [
        [CHART.left + WIDTH * 0.08, CHART.bottom - HEIGHT * 0.25],
        [CHART.left + WIDTH * 0.26, CHART.bottom - HEIGHT * 0.62],
        [CHART.left + WIDTH * 0.44, CHART.bottom - HEIGHT * 0.4],
        [CHART.left + WIDTH * 0.62, CHART.bottom - HEIGHT * 0.78],
        [CHART.left + WIDTH * 0.8, CHART.bottom - HEIGHT * 0.52],
        [CHART.left + WIDTH * 0.92, CHART.bottom - HEIGHT * 0.68]
    ];
    const path = smoothPath(points);
    return wrapSvg(`${AXIS}<path d="${path}" fill="none" stroke="${COLORS.blue}" stroke-width="2.1" stroke-linecap="round"/>${withMarkers ? markers(points, COLORS.orange, 2.6) : ''}`);
}

function scatterLine(withMarkers) {
    const points = linePoints([0.25, 0.6, 0.4, 0.8, 0.55, 0.7]);
    return wrapSvg(`${AXIS}${polyline(points, COLORS.blue, 2.2)}${withMarkers ? markers(points, COLORS.orange, 2.6) : ''}`);
}

function scatterBubble() {
    return wrapSvg(`${AXIS}
        <circle cx="${x(0.2)}" cy="${y(0.35)}" r="6" fill="${COLORS.blue}" opacity="0.7"/>
        <circle cx="${x(0.42)}" cy="${y(0.55)}" r="4.5" fill="${COLORS.orange}" opacity="0.7"/>
        <circle cx="${x(0.62)}" cy="${y(0.75)}" r="7" fill="${COLORS.green}" opacity="0.7"/>
        <circle cx="${x(0.85)}" cy="${y(0.45)}" r="5.5" fill="${COLORS.gray}" opacity="0.7"/>
    `);
}

function pieBasic() {
    return wrapSvg(`${ringSegments(20, 40, [0.4, 0.3, 0.3], [COLORS.blue, COLORS.orange, COLORS.gray])}`);
}

function pieDoughnut() {
    const segments = ringSegments(30, 18, [0.45, 0.25, 0.3], [COLORS.blue, COLORS.orange, COLORS.green]);
    const hole = '<circle cx="60" cy="40" r="11" fill="#ffffff"/>';
    return wrapSvg(`${segments}${hole}`);
}

function stockHLC(withVolume) {
    const xs = [0.1, 0.26, 0.42, 0.58, 0.74, 0.9];
    const highs = [0.85, 0.7, 0.8, 0.65, 0.75, 0.88];
    const lows = [0.35, 0.25, 0.3, 0.2, 0.28, 0.4];
    const closes = [0.5, 0.45, 0.55, 0.4, 0.6, 0.55];
    const tickLen = 3;
    const stroke = 1.2;
    const lineColor = '#595959';

    const volumeRatio = withVolume ? 0.3 : 0;
    const gapRatio = withVolume ? 0.05 : 0;
    const stockBottom = CHART.bottom - HEIGHT * volumeRatio - HEIGHT * gapRatio;
    const stockTop = CHART.top;
    const stockHeight = stockBottom - stockTop;
    const mapStock = (value) => stockBottom - value * stockHeight;

    const bars = xs.map((pct, idx) => {
        const cx = CHART.left + pct * WIDTH;
        const highY = mapStock(highs[idx]);
        const lowY = mapStock(lows[idx]);
        const closeY = mapStock(closes[idx]);
        const hlLine = `<line x1="${fmt(cx)}" y1="${fmt(highY)}" x2="${fmt(cx)}" y2="${fmt(lowY)}" stroke="${lineColor}" stroke-width="${stroke}"/>`;
        const closeTick = `<line x1="${fmt(cx)}" y1="${fmt(closeY)}" x2="${fmt(cx + tickLen)}" y2="${fmt(closeY)}" stroke="${lineColor}" stroke-width="${stroke}"/>`;
        return `${hlLine}${closeTick}`;
    }).join('');

    let volumeBars = '';
    let volAxis = '';
    if (withVolume) {
        const volTop = stockBottom + HEIGHT * gapRatio;
        const volBottom = CHART.bottom;
        const volHeight = volBottom - volTop;
        const volValues = [0.6, 0.8, 0.5, 0.7, 0.55, 0.65];
        const barW = WIDTH * 0.08;
        const tickW = 3;
        const tickCount = 3;
        let ticks = '';
        for (let i = 0; i <= tickCount; i++) {
            const ty = CHART.top + ((CHART.bottom - CHART.top) * i) / tickCount;
            ticks += `<line x1="${CHART.right - tickW}" y1="${fmt(ty)}" x2="${CHART.right + tickW}" y2="${fmt(ty)}" stroke="#c6cbd3" stroke-width="1"/>`;
        }
        volAxis = `<line x1="${CHART.right}" y1="${fmt(CHART.top)}" x2="${CHART.right}" y2="${fmt(CHART.bottom)}" stroke="#c6cbd3" stroke-width="1"/>${ticks}`;
        volumeBars = volValues.map((val, idx) => {
            const cx = CHART.left + xs[idx] * WIDTH;
            const barH = val * volHeight;
            const barY = volBottom - barH;
            return `<rect x="${fmt(cx - barW / 2)}" y="${fmt(barY)}" width="${fmt(barW)}" height="${fmt(barH)}" fill="${COLORS.blue}"/>`;
        }).join('');
    }

    return wrapSvg(`${AXIS}${bars}${volAxis}${volumeBars}`);
}

function stockOHLC(withVolume) {
    const xs = [0.12, 0.32, 0.52, 0.72, 0.92];
    const highs = [0.9, 0.75, 0.85, 0.7, 0.82];
    const lows = [0.3, 0.2, 0.25, 0.15, 0.28];
    const opens = [0.7, 0.5, 0.6, 0.45, 0.55];
    const closes = [0.45, 0.35, 0.5, 0.55, 0.7];
    const candleW = 8;
    const stroke = 1;

    const volumeRatio = withVolume ? 0.3 : 0;
    const gapRatio = withVolume ? 0.05 : 0;
    const stockBottom = CHART.bottom - HEIGHT * volumeRatio - HEIGHT * gapRatio;
    const stockTop = CHART.top;
    const stockHeight = stockBottom - stockTop;
    const mapStock = (value) => stockBottom - value * stockHeight;

    const candles = xs.map((pct, idx) => {
        const cx = CHART.left + pct * WIDTH;
        const highY = mapStock(highs[idx]);
        const lowY = mapStock(lows[idx]);
        const openY = mapStock(opens[idx]);
        const closeY = mapStock(closes[idx]);
        const isUp = closes[idx] >= opens[idx];
        const bodyTop = Math.min(openY, closeY);
        const bodyH = Math.abs(closeY - openY);
        const fill = isUp ? '#ffffff' : '#595959';
        const strokeColor = '#595959';

        const upperWick = `<line x1="${fmt(cx)}" y1="${fmt(highY)}" x2="${fmt(cx)}" y2="${fmt(bodyTop)}" stroke="${strokeColor}" stroke-width="${stroke}"/>`;
        const lowerWick = `<line x1="${fmt(cx)}" y1="${fmt(bodyTop + bodyH)}" x2="${fmt(cx)}" y2="${fmt(lowY)}" stroke="${strokeColor}" stroke-width="${stroke}"/>`;
        const body = `<rect x="${fmt(cx - candleW / 2)}" y="${fmt(bodyTop)}" width="${candleW}" height="${fmt(bodyH)}" fill="${fill}" stroke="${strokeColor}" stroke-width="${stroke}"/>`;
        return `${upperWick}${lowerWick}${body}`;
    }).join('');

    let volumeBars = '';
    let volAxis = '';
    if (withVolume) {
        const volTop = stockBottom + HEIGHT * gapRatio;
        const volBottom = CHART.bottom;
        const volHeight = volBottom - volTop;
        const volValues = [0.7, 0.9, 0.5, 0.75, 0.6];
        const barW = WIDTH * 0.1;
        const tickW = 3;
        const tickCount = 3;
        let ticks = '';
        for (let i = 0; i <= tickCount; i++) {
            const ty = CHART.top + ((CHART.bottom - CHART.top) * i) / tickCount;
            ticks += `<line x1="${CHART.right - tickW}" y1="${fmt(ty)}" x2="${CHART.right + tickW}" y2="${fmt(ty)}" stroke="#c6cbd3" stroke-width="1"/>`;
        }
        volAxis = `<line x1="${CHART.right}" y1="${fmt(CHART.top)}" x2="${CHART.right}" y2="${fmt(CHART.bottom)}" stroke="#c6cbd3" stroke-width="1"/>${ticks}`;
        volumeBars = volValues.map((val, idx) => {
            const cx = CHART.left + xs[idx] * WIDTH;
            const barH = val * volHeight;
            const barY = volBottom - barH;
            return `<rect x="${fmt(cx - barW / 2)}" y="${fmt(barY)}" width="${fmt(barW)}" height="${fmt(barH)}" fill="${COLORS.blue}"/>`;
        }).join('');
    }

    return wrapSvg(`${AXIS}${candles}${volAxis}${volumeBars}`);
}

function radarBasic() {
    const grid = radarGrid();
    const primary = radarShape(32, [0.92, 0.7, 0.85, 0.6, 0.78]);
    const secondary = radarShape(32, [0.75, 0.9, 0.68, 0.82, 0.65]);
    const primaryShape = `<polygon points="${toPoints(primary)}" fill="none" stroke="${COLORS.blue}" stroke-width="2.1"/>`;
    const secondaryShape = `<polygon points="${toPoints(secondary)}" fill="none" stroke="${COLORS.orange}" stroke-width="1.6" opacity="0.8"/>`;
    return wrapSvg(`${grid}${secondaryShape}${primaryShape}`);
}

function radarMarker() {
    const grid = radarGrid();
    const pts = radarShape(32, [1.0, 0.82, 0.9, 0.68, 0.78]);
    const secondary = radarShape(32, [0.7, 0.88, 0.6, 0.8, 0.66]);
    const secondaryShape = `<polygon points="${toPoints(secondary)}" fill="none" stroke="${COLORS.gray}" stroke-width="1.4" opacity="0.8"/>`;
    const shape = `<polygon points="${toPoints(pts)}" fill="none" stroke="${COLORS.blue}" stroke-width="2.1"/>`;
    return wrapSvg(`${grid}${secondaryShape}${shape}${markers(pts, COLORS.orange, 2.6)}`);
}

function radarFilled() {
    const grid = radarGrid();
    const pts = radarShape(32, [0.92, 0.7, 0.82, 0.6, 0.76]);
    const secondary = radarShape(32, [0.68, 0.86, 0.7, 0.84, 0.62]);
    const secondaryShape = `<polygon points="${toPoints(secondary)}" fill="none" stroke="${COLORS.orange}" stroke-width="1.6" opacity="0.8"/>`;
    const shape = `<polygon points="${toPoints(pts)}" fill="${COLORS.blue}" opacity="0.3" stroke="${COLORS.blue}" stroke-width="2"/>`;
    return wrapSvg(`${grid}${secondaryShape}${shape}`);
}

function sunburstBasic() {
    const outer = ringSegments(32, 8, [0.5, 0.25, 0.25], [COLORS.blue, COLORS.orange, COLORS.gray]);
    const middle = ringSegments(22, 8, [0.35, 0.35, 0.3], [COLORS.green, COLORS.yellow, COLORS.lightBlue]);
    const inner = ringSegments(12, 8, [0.6, 0.4], [COLORS.purple, COLORS.blue]);
    return wrapSvg(`${outer}${middle}${inner}`);
}

function treemapBasic() {
    const leftWidth = WIDTH * 0.62;
    const rightWidth = WIDTH - leftWidth;
    const topHeight = HEIGHT * 0.62;
    const bottomHeight = HEIGHT - topHeight;
    const leftSplit = leftWidth * 0.58;
    const rightTop = topHeight * 0.55;
    return wrapSvg(`
        <rect x="${fmt(CHART.left)}" y="${fmt(CHART.top)}" width="${fmt(leftWidth)}" height="${fmt(topHeight)}" fill="${COLORS.blue}" stroke="#ffffff" stroke-width="1"/>
        <rect x="${fmt(CHART.left)}" y="${fmt(CHART.top + topHeight)}" width="${fmt(leftSplit)}" height="${fmt(bottomHeight)}" fill="${COLORS.yellow}" stroke="#ffffff" stroke-width="1"/>
        <rect x="${fmt(CHART.left + leftSplit)}" y="${fmt(CHART.top + topHeight)}" width="${fmt(leftWidth - leftSplit)}" height="${fmt(bottomHeight)}" fill="${COLORS.green}" stroke="#ffffff" stroke-width="1"/>
        <rect x="${fmt(CHART.left + leftWidth)}" y="${fmt(CHART.top)}" width="${fmt(rightWidth)}" height="${fmt(rightTop)}" fill="${COLORS.orange}" stroke="#ffffff" stroke-width="1"/>
        <rect x="${fmt(CHART.left + leftWidth)}" y="${fmt(CHART.top + rightTop)}" width="${fmt(rightWidth)}" height="${fmt(topHeight - rightTop)}" fill="${COLORS.gray}" stroke="#ffffff" stroke-width="1"/>
        <rect x="${fmt(CHART.left + leftWidth)}" y="${fmt(CHART.top + topHeight)}" width="${fmt(rightWidth)}" height="${fmt(bottomHeight)}" fill="${COLORS.lightBlue}" stroke="#ffffff" stroke-width="1"/>
    `);
}

function funnelBasic() {
    const center = 0.5;
    const centerX = CHART.left + center * WIDTH;
    const segments = [
        { top: 0.94, bottom: 0.8, wTop: 0.92, wBottom: 0.76, color: COLORS.blue },
        { top: 0.74, bottom: 0.6, wTop: 0.74, wBottom: 0.6, color: COLORS.orange },
        { top: 0.54, bottom: 0.4, wTop: 0.56, wBottom: 0.44, color: COLORS.gray },
        { top: 0.34, bottom: 0.2, wTop: 0.4, wBottom: 0.28, color: COLORS.green }
    ];
    const shapes = segments.map((seg) => {
        const topY = y(seg.top);
        const bottomY = y(seg.bottom);
        const halfTop = (seg.wTop * WIDTH) / 2;
        const halfBottom = (seg.wBottom * WIDTH) / 2;
        const leftTop = centerX - halfTop;
        const rightTop = centerX + halfTop;
        const leftBottom = centerX - halfBottom;
        const rightBottom = centerX + halfBottom;
        return `<polygon points="${fmt(leftTop)},${topY} ${fmt(rightTop)},${topY} ${fmt(rightBottom)},${bottomY} ${fmt(leftBottom)},${bottomY}" fill="${seg.color}" stroke="#ffffff" stroke-width="1"/>`;
    }).join('');
    return wrapSvg(shapes);
}

function waterfallBasic() {
    const steps = [0.45, -0.22, 0.32];
    const positions = [0.14, 0.4, 0.66, 0.88];
    const widthPct = 0.12;
    let cumulative = 0;
    const bars = [];
    const connectors = [];
    steps.forEach((delta, idx) => {
        const start = cumulative;
        cumulative = Math.max(0, Math.min(0.9, cumulative + delta));
        const color = delta >= 0 ? COLORS.green : COLORS.orange;
        bars.push(floatingBar(positions[idx], widthPct, start, cumulative, color));
        const endY = y(cumulative);
        const xRight = CHART.left + positions[idx] * WIDTH + (WIDTH * widthPct) / 2;
        const xLeftNext = CHART.left + positions[idx + 1] * WIDTH - (WIDTH * widthPct) / 2;
        connectors.push(`<line x1="${fmt(xRight)}" y1="${endY}" x2="${fmt(xLeftNext)}" y2="${endY}" stroke="#9099a8" stroke-width="1"/>`);
    });
    const totalBar = floatingBar(positions[3], widthPct, 0, cumulative, COLORS.blue);
    return wrapSvg(`${AXIS}${bars.join('')}${connectors.join('')}${totalBar}`);
}

function comboBarLine(secondary) {
    const bars = `
        ${bar(0.18, 0.08, 0.52, COLORS.blue)}
        ${bar(0.38, 0.08, 0.7, COLORS.blue)}
        ${bar(0.58, 0.08, 0.38, COLORS.blue)}
        ${bar(0.78, 0.08, 0.78, COLORS.blue)}
    `;
    const line = polyline(linePoints(LINE_VALUES), secondary ? COLORS.orange : COLORS.green, 2.2);
    return wrapSvg(`${AXIS}${bars}${line}`);
}

function comboAreaBar() {
    const bars = `
        ${bar(0.2, 0.08, 0.48, COLORS.orange)}
        ${bar(0.4, 0.08, 0.68, COLORS.orange)}
        ${bar(0.6, 0.08, 0.38, COLORS.orange)}
        ${bar(0.8, 0.08, 0.76, COLORS.orange)}
    `;
    const area = areaFill(linePoints(LINE_VALUES), COLORS.blue, 0.2);
    return wrapSvg(`${AXIS}${area}${bars}${polyline(linePoints(LINE_VALUES), COLORS.blue, 2.2)}`);
}

const CHART_SVG_BUILDERS = {
    bar_clustered: barClustered,
    bar_stacked: barStacked,
    bar_percent: barPercent,
    histogram_basic: histogramBasic,
    barh_clustered: barHClustered,
    barh_stacked: barHStacked,
    barh_percent: barHPercent,
    line_basic: () => lineBasic(false),
    line_stacked: () => lineStacked(false),
    line_percent: () => linePercent(false),
    line_marker: () => lineBasic(true),
    line_marker_stacked: () => lineStacked(true),
    line_marker_percent: () => linePercent(true),
    pie_basic: pieBasic,
    pie_doughnut: pieDoughnut,
    area_basic: areaBasic,
    area_stacked: areaStacked,
    area_percent: areaPercent,
    scatter_basic: scatterBasic,
    scatter_smooth_marker: () => scatterSmooth(true),
    scatter_smooth: () => scatterSmooth(false),
    scatter_line_marker: () => scatterLine(true),
    scatter_line: () => scatterLine(false),
    scatter_bubble: scatterBubble,
    stock_hlc: () => stockHLC(false),
    stock_ohlc: () => stockOHLC(false),
    stock_vol_hlc: () => stockHLC(true),
    stock_vol_ohlc: () => stockOHLC(true),
    radar_basic: radarBasic,
    radar_marker: radarMarker,
    radar_filled: radarFilled,
    sunburst_basic: sunburstBasic,
    treemap_basic: treemapBasic,
    funnel_basic: funnelBasic,
    waterfall_basic: waterfallBasic,
    combo_bar_line: () => comboBarLine(false),
    combo_area_bar: comboAreaBar
};

export function getChartSvg(key) {
    const builder = CHART_SVG_BUILDERS[key];
    return builder ? builder() : '';
}
