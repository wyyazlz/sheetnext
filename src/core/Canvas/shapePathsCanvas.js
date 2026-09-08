import presets from '../Drawing/presetShapeData.js';

const ANGLE = Math.PI / 10800000;
const TURN = Math.PI * 2;

function evaluateGuide([, operation, firstToken, secondToken, thirdToken], resolve) {
    const first = resolve(firstToken), second = resolve(secondToken), third = resolve(thirdToken);
    switch (operation) {
        case 'val': return first;
        case '+-': return first + second - third;
        case '*/': return third === 0 ? 0 : first * second / third;
        case '+/': return third === 0 ? 0 : (first + second) / third;
        case '?:': return first > 0 ? second : third;
        case 'abs': return Math.abs(first);
        case 'min': return Math.min(first, second);
        case 'max': return Math.max(first, second);
        case 'sqrt': return Math.sqrt(Math.max(0, first));
        case 'mod': return Math.hypot(first, second, third);
        case 'pin': return Math.max(first, Math.min(second, third));
        case 'at2': return Math.atan2(second, first) / ANGLE;
        case 'sin': return first * Math.sin(second * ANGLE);
        case 'cos': return first * Math.cos(second * ANGLE);
        case 'tan': return first * Math.tan(second * ANGLE);
        case 'cat2': return first * Math.cos(Math.atan2(third, second));
        case 'sat2': return first * Math.sin(Math.atan2(third, second));
        default: return 0;
    }
}

function createGuides(preset, width, height, adjustments) {
    const values = Object.assign(Object.create(null), {
        l: 0, t: 0, r: width, b: height, w: width, h: height,
        hc: width / 2, vc: height / 2, ss: Math.min(width, height), ls: Math.max(width, height),
        cd2: 10800000, cd4: 5400000, cd8: 2700000,
        '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000
    });
    const resolve = token => {
        if (typeof token === 'number') return token;
        if (token === undefined) return 0;
        if (Object.hasOwn(values, token)) return values[token];
        const division = /^(w|h|ss)d(\d+)$/.exec(token);
        if (division) return values[division[1]] / Number(division[2]);
        return Number(token) || 0;
    };
    for (const guide of preset.adjustments) {
        const override = adjustments?.[guide[0]];
        values[guide[0]] = typeof override === 'number' && Number.isFinite(override)
            ? override : evaluateGuide(guide, resolve);
    }
    for (const guide of preset.guides) {
        const value = evaluateGuide(guide, resolve);
        values[guide[0]] = Number.isFinite(value) ? value : 0;
    }
    return resolve;
}

function addPoint(bounds, pointX, pointY) {
    bounds.left = Math.min(bounds.left, pointX);
    bounds.top = Math.min(bounds.top, pointY);
    bounds.right = Math.max(bounds.right, pointX);
    bounds.bottom = Math.max(bounds.bottom, pointY);
}

/** DrawingML angles are radial angles; Canvas ellipses require parametric angles. */
function ellipseAngle(angle, rx, ry) {
    return Math.atan2(rx * Math.sin(angle), ry * Math.cos(angle));
}

/** @param {Object} definition @param {Function} resolve @param {number} width @param {number} height @param {Object} bounds */
function createPath(definition, resolve, width, height, bounds) {
    const path = new Path2D();
    const sx = definition.width ? width / definition.width : 1;
    const sy = definition.height ? height / definition.height : 1;
    let px = 0, py = 0, startX = 0, startY = 0;

    for (const [type, ...tokens] of definition.commands) {
        const args = tokens.map(resolve);
        if (type === 'Z') {
            path.closePath();
            px = startX;
            py = startY;
        } else if (type === 'A') {
            const [radiusX, radiusY, start, sweep] = args;
            const rx = Math.abs(radiusX), ry = Math.abs(radiusY);
            if (rx === 0 || ry === 0 || sweep === 0) continue;
            const startAngle = ellipseAngle(start * ANGLE, rx, ry);
            let delta = ellipseAngle((start + sweep) * ANGLE, rx, ry) - startAngle;
            if (sweep > 0 && delta < 0) delta += TURN;
            if (sweep < 0 && delta > 0) delta -= TURN;
            if (Math.abs(sweep) >= 21600000) delta = Math.sign(sweep) * TURN;
            const cx = px - rx * Math.cos(startAngle);
            const cy = py - ry * Math.sin(startAngle);
            const endAngle = startAngle + delta;
            path.ellipse(cx * sx, cy * sy, rx * sx, ry * sy, 0, startAngle, endAngle, delta < 0);
            px = cx + rx * Math.cos(endAngle);
            py = cy + ry * Math.sin(endAngle);
            addPoint(bounds, px * sx, py * sy);
            for (let quadrant = 0; quadrant < 4; quadrant++) {
                const angle = quadrant * Math.PI / 2;
                const distance = ((delta < 0 ? startAngle - angle : angle - startAngle) % TURN + TURN) % TURN;
                if (distance <= Math.abs(delta) + 1e-10) {
                    addPoint(bounds, (cx + rx * Math.cos(angle)) * sx, (cy + ry * Math.sin(angle)) * sy);
                }
            }
        } else {
            const coords = args.map((value, index) => value * (index % 2 ? sy : sx));
            for (let index = 0; index < coords.length; index += 2) addPoint(bounds, coords[index], coords[index + 1]);
            if (type === 'M') {
                path.moveTo(...coords);
                startX = args[0];
                startY = args[1];
            } else if (type === 'L') path.lineTo(...coords);
            else if (type === 'Q') path.quadraticCurveTo(...coords);
            else if (type === 'C') path.bezierCurveTo(...coords);
            px = args[args.length - 2];
            py = args[args.length - 1];
        }
    }
    return { path, fill: definition.fill, stroke: definition.stroke };
}

/** Resolve preset geometry once per drawing, size and adjustment set. Coordinates are local to the shape. */
export function getShapeGeometry(drawing, width, height) {
    const shapeType = drawing.shapeType || 'rect';
    const adjustments = drawing.shapeAdjustments || {};
    const key = JSON.stringify([shapeType, width, height, adjustments]);
    if (drawing._shapeGeometry?.key === key) return drawing._shapeGeometry;
    const preset = Object.hasOwn(presets, shapeType) ? presets[shapeType] : presets.rect;
    const resolve = createGuides(preset, width, height, adjustments);
    const bounds = { left: 0, top: 0, right: width, bottom: height };
    const paths = preset.paths.map(path => createPath(path, resolve, width, height, bounds));
    const [left, top, right, bottom] = preset.rect?.map(resolve) || [0, 0, width, height];
    drawing._shapeGeometry = {
        key, paths, bounds,
        textRect: { x: left, y: top, w: Math.max(0, right - left), h: Math.max(0, bottom - top) }
    };
    return drawing._shapeGeometry;
}

/** Include callout leaders and stroke joins without changing the shape's anchor rectangle. */
export function getShapeRenderBounds(drawing, width, height) {
    const bounds = getShapeGeometry(drawing, width, height).bounds;
    const padding = Math.max(2, Math.ceil((Number(drawing.shapeStyle?.strokeWidth) || 0) * 5));
    const left = Math.floor(bounds.left - padding);
    const top = Math.floor(bounds.top - padding);
    return {
        x: left, y: top,
        w: Math.ceil(bounds.right + padding) - left,
        h: Math.ceil(bounds.bottom + padding) - top
    };
}
