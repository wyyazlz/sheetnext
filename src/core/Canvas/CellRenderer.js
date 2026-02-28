import { interpolateColor } from './helpers.js';
import { shouldHideFormula } from './helpers.js'

/**
 * 单元格渲染模块
 * 负责处理单元格的渲染逻辑,包括背景、文本、边框等
 */

/**
 * 获取单元格的超级表样式（如果有）
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} col
 * @returns {Object|null} { bgColor, fgColor, isBold } 或 null
 */
function getAutoTableFgColor(bgColor) {
    if (typeof bgColor !== 'string') return null;

    const matched = bgColor.trim().match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (!matched) return null;

    let hex = matched[1];
    if (hex.length === 3) {
        hex = hex.split('').map(ch => ch + ch).join('');
    }

    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    const luminance = (r * 299 + g * 587 + b * 114) / 1000;

    return luminance < 150 ? '#FFFFFF' : null;
}

function pickDefined(...values) {
    for (const value of values) {
        if (value !== undefined) return value;
    }
    return undefined;
}

function mergeTableStyleWithDxf(baseStyle, dxfStyle) {
    if (!dxfStyle) return baseStyle;

    const nextStyle = { ...baseStyle };
    if (dxfStyle.fill?.fgColor) {
        nextStyle.bgColor = dxfStyle.fill.fgColor;
    }
    if (dxfStyle.font?.color) {
        nextStyle.fgColor = dxfStyle.font.color;
    }
    if (typeof dxfStyle.font?.bold === 'boolean') {
        nextStyle.isBold = dxfStyle.font.bold;
    }

    return nextStyle;
}

function finalizeTableCellStyle(style, dataFallbackFg) {
    const bgColor = style.bgColor ?? null;
    const fgColor = style.fgColor ?? dataFallbackFg ?? getAutoTableFgColor(bgColor);
    return {
        bgColor,
        fgColor,
        isBold: !!style.isBold,
        borderColor: style.borderColor ?? null,
        _fromImport: style._fromImport === true
    };
}

function getStripeColor(index, stripe1Color, stripe2Color, stripe1Size = 1, stripe2Size = 1) {
    const s1 = Math.max(1, stripe1Size || 1);
    const s2 = Math.max(1, stripe2Size || 1);
    const cycle = s1 + s2;
    const offset = index % cycle;
    return offset < s1 ? stripe1Color : stripe2Color;
}

function applyTableBorderColor(border, color) {
    if (!border || typeof border !== 'object' || !color) return border || {};

    let changed = false;
    const next = { ...border };
    const sides = ['top', 'right', 'bottom', 'left'];

    for (const side of sides) {
        const item = border[side];
        if (!item?.style) continue;
        next[side] = { ...item, color };
        changed = true;
    }

    if (border.diagonal?.style) {
        next.diagonal = { ...border.diagonal, color };
        changed = true;
    }

    return changed ? next : border;
}

function getTableCellStyle(sheet, row, col) {
    if (!sheet?.Table?.size) return null;

    const table = sheet.Table.getTableAt(row, col);
    if (!table) return null;

    const style = table.style;
    if (!style) return null;
    const allowDxfOverlay = style._fromTokens !== true;

    const dataFallbackFg = style.dataFg;

    if (table.isHeaderCell(row, col)) {
        const headerStyle = {
            bgColor: style.headerBg,
            fgColor: style.headerFg ?? dataFallbackFg,
            isBold: style.headerBold !== false,
            borderColor: style.headerBorder ?? style.border,
            _fromImport: table._fromImport === true
        };
        const merged = allowDxfOverlay
            ? mergeTableStyleWithDxf(headerStyle, table._getHeaderDxfStyle?.())
            : headerStyle;
        return finalizeTableCellStyle(merged, dataFallbackFg);
    }

    if (table.isTotalsCell(row, col)) {
        const totalsStyle = {
            bgColor: style.totalsBg,
            fgColor: style.totalsFg ?? dataFallbackFg,
            isBold: style.totalsBold !== false,
            borderColor: style.border,
            _fromImport: table._fromImport === true
        };
        return finalizeTableCellStyle(totalsStyle, dataFallbackFg);
    }

    if (table.isDataCell(row, col)) {
        const dataRowIndex = table.getDataRowIndex(row);
        const colIndex = table.getColumnIndex(col);

        let bgColor = style.dataBg ?? null;
        let fgColor = dataFallbackFg ?? null;
        let isBold = !!style.dataBold;

        if (table.showRowStripes && dataRowIndex >= 0) {
            bgColor = getStripeColor(
                dataRowIndex,
                style.rowStripe1,
                style.rowStripe2,
                style.rowStripe1Size,
                style.rowStripe2Size
            );
        }

        if (table.showColumnStripes && colIndex >= 0) {
            const stripe1Color = pickDefined(style.columnStripe1, style.rowStripe2, style.rowStripe1, bgColor);
            const stripe2Color = pickDefined(style.columnStripe2, style.rowStripe1, style.rowStripe2, bgColor);
            bgColor = getStripeColor(
                colIndex,
                stripe1Color,
                stripe2Color,
                style.columnStripe1Size,
                style.columnStripe2Size
            );
        }

        if (table.showFirstColumn && colIndex === 0) {
            bgColor = pickDefined(style.firstColumnBg, bgColor);
            fgColor = pickDefined(style.firstColumnFg, fgColor);
            isBold = style.firstColumnBold ?? isBold;
        }

        const lastColIndex = table.range.e.c - table.range.s.c;
        if (table.showLastColumn && colIndex === lastColIndex) {
            bgColor = pickDefined(style.lastColumnBg, bgColor);
            fgColor = pickDefined(style.lastColumnFg, fgColor);
            isBold = style.lastColumnBold ?? isBold;
        }

        const dataStyle = allowDxfOverlay
            ? mergeTableStyleWithDxf(
                { bgColor, fgColor, isBold, borderColor: style.border, _fromImport: table._fromImport === true },
                table._getDataDxfStyle?.(colIndex)
            )
            : { bgColor, fgColor, isBold, borderColor: style.border, _fromImport: table._fromImport === true };

        return finalizeTableCellStyle(dataStyle, dataFallbackFg);
    }

    return null;
}

export function rCellBackground(x, y, w, h, cell, skipGridLine = false, cfFormat = null, tableStyle = null) {
    // 条件格式背景色优先
    const cfFillColor = cfFormat?.fill?.fgColor;
    const preferImportedTableStyle = tableStyle?._fromImport === true;

    // 确定最终背景色（优先级：条件格式 > 单元格填充 > 超级表样式）
    let fillColor = null;

    if (cfFillColor) {
        fillColor = cfFillColor;
    } else if (preferImportedTableStyle && tableStyle?.bgColor) {
        fillColor = tableStyle.bgColor;
    } else if (cell.fill.type === 'pattern' && cell.fill.pattern === 'solid' && cell.fill.fgColor) {
        fillColor = cell.fill.fgColor;
    } else if (tableStyle?.bgColor) {
        fillColor = tableStyle.bgColor;
    }

    if (fillColor) {
        this.bc.fillStyle = fillColor;
        this.bc.fillRect(x, y, w, h);
    } else if (cell.fill.type === 'pattern' && (cell.fill.pattern == 'solid' || cell.fill.pattern == 'none')) {
        const defaultBg = this.eyeProtectionMode ? this.eyeProtectionColor : '#ffffff';
        this.bc.fillStyle = cell.fill.fgColor ?? defaultBg;
        this.bc.fillRect(x, y, w, h);
    } else if (cell.fill.type === 'pattern') { // 填充底纹
        // 清空临时画布,准备重新绘制
        this.pCtx.clearRect(0, 0, 8, 8);
        // 先填充背景色
        const patternDefaultBg = this.eyeProtectionMode ? this.eyeProtectionColor : '#ffffff';
        this.pCtx.fillStyle = cell.fill.bgColor || patternDefaultBg;
        this.pCtx.fillRect(0, 0, 8, 8);
        // 设置前景色
        this.pCtx.fillStyle = cell.fill.fgColor || '#000000';
        this.pCtx.strokeStyle = cell.fill.fgColor || '#000000';
        // 根据 pattern 名称绘制对应的 18 种底纹
        switch (cell.fill.pattern) {
            case 'solid':
                // 纯色：直接用前景色填满整个 8×8
                this.pCtx.fillRect(0, 0, 8, 8);
                break;
            case 'darkGray':
                // 50% 灰度：纵向两列、水平两行交替填充
                this.pCtx.fillRect(0, 0, 4, 4);
                this.pCtx.fillRect(4, 4, 4, 4);
                break;
            case 'mediumGray':
                // 75% 灰度：三个 8×8 中的像素块为前景色（留一个像素空白）
                this.pCtx.fillRect(0, 0, 8, 4);
                this.pCtx.fillRect(0, 4, 4, 4);
                break;
            case 'lightGray':
                // 25% 灰度：只画四个角的小点
                this.pCtx.fillRect(0, 0, 2, 2);
                this.pCtx.fillRect(6, 0, 2, 2);
                this.pCtx.fillRect(0, 6, 2, 2);
                this.pCtx.fillRect(6, 6, 2, 2);
                break;
            case 'gray125':
                // 12.5% 灰度：在 8×8 区域里画两个单像素点,分布在 (0,0) 和 (4,4)
                this.pCtx.fillRect(0, 0, 1, 1);
                this.pCtx.fillRect(4, 4, 1, 1);
                break;
            case 'gray0625':
                // 6.25% 灰度：在 8×8 区域里只画一个单像素点
                this.pCtx.fillRect(0, 0, 1, 1);
                break;
            case 'darkHorizontal':
                // 深色水平线：每隔 4 像素画一条横线
                for (let y0 = 0; y0 < 8; y0 += 4) {
                    this.pCtx.fillRect(0, y0, 8, 2);
                }
                break;
            case 'darkVertical':
                // 深色垂直线：每隔 4 像素画一条竖线
                for (let x0 = 0; x0 < 8; x0 += 4) {
                    this.pCtx.fillRect(x0, 0, 2, 8);
                }
                break;
            case 'darkDown':
                // 深色向下对角线：从左上到右下画像素点
                for (let i = -8; i < 16; i += 4) {
                    this.pCtx.beginPath();
                    this.pCtx.moveTo(i, 0);
                    this.pCtx.lineTo(i + 8, 8);
                    this.pCtx.stroke();
                }
                break;
            case 'darkUp':
                // 深色向上对角线：从左下到右上画像素点
                for (let i = 0; i < 16; i += 4) {
                    this.pCtx.beginPath();
                    this.pCtx.moveTo(i, 8);
                    this.pCtx.lineTo(i - 8, 0);
                    this.pCtx.stroke();
                }
                break;
            case 'darkGrid':
                // 深色网格：水平线 & 垂直线
                for (let y0 = 0; y0 < 8; y0 += 4) {
                    this.pCtx.fillRect(0, y0, 8, 1);
                }
                for (let x0 = 0; x0 < 8; x0 += 4) {
                    this.pCtx.fillRect(x0, 0, 1, 8);
                }
                break;
            case 'darkTrellis':
                // 深色格子：在 darkGrid 基础上加对角线
                // 先画水平 & 垂直
                for (let y0 = 0; y0 < 8; y0 += 4) {
                    this.pCtx.fillRect(0, y0, 8, 1);
                }
                for (let x0 = 0; x0 < 8; x0 += 4) {
                    this.pCtx.fillRect(x0, 0, 1, 8);
                }
                // 再画两条对角线
                this.pCtx.beginPath();
                this.pCtx.moveTo(0, 0);
                this.pCtx.lineTo(8, 8);
                this.pCtx.stroke();
                this.pCtx.beginPath();
                this.pCtx.moveTo(8, 0);
                this.pCtx.lineTo(0, 8);
                this.pCtx.stroke();
                break;
            case 'lightHorizontal':
                // 浅色水平线：每隔 8 像素画一条 1 像素高的横线
                this.pCtx.fillRect(0, 0, 8, 1);
                break;
            case 'lightVertical':
                // 浅色垂直线：每隔 8 像素画一条 1 像素宽的竖线
                this.pCtx.fillRect(0, 0, 1, 8);
                break;
            case 'lightDown':
                // 浅色向下对角线（稀疏版）：一条对角线
                this.pCtx.beginPath();
                this.pCtx.moveTo(0, 0);
                this.pCtx.lineTo(8, 8);
                this.pCtx.stroke();
                break;
            case 'lightUp':
                // 浅色向上对角线（稀疏版）：一条对角线
                this.pCtx.beginPath();
                this.pCtx.moveTo(8, 0);
                this.pCtx.lineTo(0, 8);
                this.pCtx.stroke();
                break;
            case 'lightGrid':
                // 浅色网格：只画横线或竖线中的一部分（稀疏）
                this.pCtx.fillRect(0, 0, 8, 1);
                this.pCtx.fillRect(0, 0, 1, 8);
                break;
            case 'lightTrellis':
                // 浅色格子：一条对角线 + 一条水平或竖直线
                this.pCtx.fillRect(0, 0, 8, 1);
                this.pCtx.beginPath();
                this.pCtx.moveTo(0, 0);
                this.pCtx.lineTo(8, 8);
                this.pCtx.stroke();
                break;
            default:
                // 若遇到未知 pattern,则退回到纯色填充
                this.pCtx.fillRect(0, 0, 8, 8);
                break;
        }
        // 利用临时画布生成重复图案
        const pattern = this.bc.createPattern(this.pCanvas, 'repeat');
        this.bc.fillStyle = pattern;
        this.bc.fillRect(x, y, w, h);
    } else if (cell.fill.type === 'gradient') { // 渐变填充
        // 从 cell.fill 中读取相关参数,提供默认值以防缺失
        const {
            gradientType = 'linear', // 'linear' 或 'path'
            degree = 0,             // 线性渐变角度,单位：度,Excel 中 0° 表示从左到右
            left = 0,               // 径向渐变(inner rectangle) 相对于单元格左边界的比例 [0,1]
            right = 1,              // 径向渐变(inner rectangle) 相对于单元格左边界的比例 [0,1]
            top = 0,                // 径向渐变(inner rectangle) 相对于单元格上边界的比例 [0,1]
            bottom = 1,             // 径向渐变(inner rectangle) 相对于单元格上边界的比例 [0,1]
            stops = []              // 渐变停靠点数组,元素形如 { position: "0.0" ~ "1.0", color: "#RRGGBB" }
        } = cell.fill;

        const cx = x + w / 2;
        const cy = y + h / 2;

        if (gradientType === 'path') {
            // Excel的径向渐变是矩形扩散,需要自定义实现
            // 计算内矩形的实际坐标
            const innerLeft = x + w * left;
            const innerRight = x + w * right;
            const innerTop = y + h * top;
            const innerBottom = y + h * bottom;

            // 内矩形的宽高和中心
            const innerWidth = innerRight - innerLeft;
            const innerHeight = innerBottom - innerTop;
            const innerCx = (innerLeft + innerRight) / 2;
            const innerCy = (innerTop + innerBottom) / 2;

            // 创建临时canvas用于自定义渐变
            const tempCanvas = document.createElement('canvas');
            tempCanvas.width = w;
            tempCanvas.height = h;
            const tempCtx = tempCanvas.getContext('2d');

            // 计算从内矩形到外矩形的最大扩展比例
            const maxScaleX = innerWidth > 0 ? Math.max((innerCx - x) / (innerWidth / 2), (x + w - innerCx) / (innerWidth / 2)) : w;
            const maxScaleY = innerHeight > 0 ? Math.max((innerCy - y) / (innerHeight / 2), (y + h - innerCy) / (innerHeight / 2)) : h;
            const maxScale = Math.max(maxScaleX, maxScaleY);

            // 绘制多层矩形来模拟渐变
            const layers = 100; // 层数,越多越平滑

            for (let i = layers - 1; i >= 0; i--) {
                const t = i / (layers - 1); // 当前层的位置 [0, 1]

                // 根据stops计算当前层的颜色
                let color = stops[0]?.color || '#FFFFFF';
                for (let j = 0; j < stops.length - 1; j++) {
                    const stop1 = parseFloat(stops[j].position);
                    const stop2 = parseFloat(stops[j + 1].position);
                    if (t >= stop1 && t <= stop2) {
                        // 在两个色标之间插值
                        const ratio = (t - stop1) / (stop2 - stop1);
                        color = interpolateColor(stops[j].color, stops[j + 1].color, ratio);
                        break;
                    } else if (t > stop2) {
                        color = stops[j + 1].color;
                    }
                }

                // 计算当前层矩形的大小
                const scale = 1 + (maxScale - 1) * t;
                const rectWidth = Math.max(1, innerWidth * scale);
                const rectHeight = Math.max(1, innerHeight * scale);
                const rectX = innerCx - rectWidth / 2 - x;
                const rectY = innerCy - rectHeight / 2 - y;

                tempCtx.fillStyle = color;
                tempCtx.fillRect(rectX, rectY, rectWidth, rectHeight);
            }

            // 将临时canvas绘制到主canvas
            this.bc.drawImage(tempCanvas, x, y);
        } else {
            // 先确保 degree 是数字,再归一化
            const deg = ((parseFloat(degree) % 360) + 360) % 360;

            // 对于整90°倍数,直接使用边界坐标,避免 size 引入过大/过小导致渐变区间不在预期范围
            let x0, y0, x1, y1;
            if (deg === 0) {
                // 从左到右
                x0 = x; y0 = y + h / 2;
                x1 = x + w; y1 = y + h / 2;
            } else if (deg === 90) {
                // 从上到下
                x0 = x + w / 2; y0 = y;
                x1 = x + w / 2; y1 = y + h;
            } else if (deg === 180) {
                // 从右到左
                x0 = x + w; y0 = y + h / 2;
                x1 = x; y1 = y + h / 2;
            } else if (deg === 270) {
                // 从下到上
                x0 = x + w / 2; y0 = y + h;
                x1 = x + w / 2; y1 = y;
            } else {
                // 其他角度,保持"先当作正方形再缩放"的逻辑
                const rad = deg * Math.PI / 180;
                const size = Math.max(w, h);
                const halfL = size / 2;
                const dx = Math.cos(rad) * halfL;
                const dy = Math.sin(rad) * halfL;
                const cx = x + w / 2;
                const cy = y + h / 2;
                x0 = cx - dx;
                y0 = cy - dy;
                x1 = cx + dx;
                y1 = cy + dy;
            }
            // 创建并应用线性渐变
            const grad = this.bc.createLinearGradient(x0, y0, x1, y1);
            for (let stop of stops) {
                grad.addColorStop(parseFloat(stop.position), stop.color);
            }
            this.bc.fillStyle = grad;
            this.bc.fillRect(x, y, w, h);


        }

    } else {
        const elseBg = this.eyeProtectionMode ? this.eyeProtectionColor : '#ffffff';
        this.bc.fillStyle = cell.fill.fgColor ?? elseBg;
        this.bc.fillRect(x, y, w, h);
    }

    // 绘制网格线
    if (!skipGridLine && this.activeSheet.showGridLines) {
        const rect = this.snapRect(x, y, w, h);
        this.bc.save();
        this.bc.strokeStyle = cell.fill.fgColor ?? '#ddd';
        this.bc.lineWidth = this.px(1);
        this.bc.strokeRect(rect.x, rect.y, rect.w, rect.h);
        this.bc.restore();
    }
}

// 渲染单元格文本
export function rCellText(x, y, w, h, cell) {
    // 显示公式模式：显示 editVal 而不是 showVal
    const sheet = this.activeSheet;
    const hideFormula = shouldHideFormula(sheet, cell);
    const showFormula = sheet?.showFormulas && !hideFormula;
    const v = showFormula ? cell.editVal.toString() : cell.showVal.toString();
    if (!v) return;

    // 富文本路径（非公式模式且有富文本数据）
    const richText = !showFormula ? cell._richText : null;

    // 处理文本换行
    let lines = [];
    if (cell.alignment.wrapText) {
        if (richText) {
            lines = wrapRichText.call(this, richText, w - 4, cell);
        } else {
            lines = wrapText.call(this, v, w - 4, this.bc); // 4是左右边距
        }
    } else {
        if (richText) {
            lines = splitRichTextByNewline(richText);
        } else {
            lines = v.split('\n');
        }
    }

    rText.call(this, { lines, cell, x, y, w, h, richText: !!richText }, false);
}

// 渲染带溢出支持的单元格文本（支持条件格式字体色和超级表样式）
export function rCellTextWithOverflow(r, c, x, y, w, h, cell, sheet, nonEmptyCells, cfFormat = null, tableStyle = null) {
    // 显示公式模式：显示 editVal 而不是 showVal
    const hideFormula = shouldHideFormula(sheet, cell);
    const showFormula = sheet?.showFormulas && !hideFormula;
    const v = showFormula ? cell.editVal.toString() : cell.showVal.toString();
    if (!v) return;

    // 富文本路径
    const richText = !showFormula ? cell._richText : null;

    // 处理文本换行
    let lines = [];
    if (cell.alignment.wrapText) {
        if (richText) {
            lines = wrapRichText.call(this, richText, w - 4, cell);
        } else {
            lines = wrapText.call(this, v, w - 4, this.bc);
        }
        // 换行模式不支持溢出
        rText.call(this, { lines, cell, x, y, w, h, cfFormat, tableStyle, richText: !!richText }, false);
        return;
    } else {
        if (richText) {
            lines = splitRichTextByNewline(richText);
        } else {
            lines = v.split('\n');
        }
    }

    // 设置字体以测量文本宽度
    const f = cell.font;
    const fsize = (f.size ?? 11) * 1.333;
    this.bc.font = `${f.bold ? 'bold' : ''} ${f.italic ? 'italic' : ''} ${fsize}px ${f.name ?? '宋体'}`;

    // 计算第一行文本宽度（通常溢出只考虑第一行）
    let textWidth;
    if (richText && lines.length > 0 && Array.isArray(lines[0])) {
        textWidth = measureRichLineWidth.call(this, lines[0], cell);
    } else {
        textWidth = lines.length > 0 ? this.bc.measureText(lines[0]).width : 0;
    }
    const horizontalAlign = cell.horizontalAlign;

    // 如果文本未超出单元格，直接渲染
    if (textWidth <= w - 4) {
        rText.call(this, { lines, cell, x, y, w, h, cfFormat, tableStyle, richText: !!richText }, false);
        return;
    }

    // 文本超出，计算可以延伸的范围
    let extendLeft = 0;
    let extendRight = 0;
    let leftCols = [];
    let rightCols = [];

    // 根据对齐方式确定延伸方向
    if (horizontalAlign === 'left' || !horizontalAlign) {
        // 左对齐：向右延伸
        let currentCol = c + 1;
        let accumulatedWidth = w;

        while (accumulatedWidth < textWidth + 4 && currentCol < sheet.colCount) {
            const col = sheet.getCol(currentCol);

            // 跳过隐藏列
            if (col.hidden) {
                currentCol++;
                continue;
            }

            const key = `${r},${currentCol}`;
            const cellInfo = sheet.getCell(r, currentCol);

            // 如果相邻单元格非空、已被占用、或是合并单元格，停止延伸
            if (nonEmptyCells.has(key) || this.hiddenBorders.has(key) || cellInfo.isMerged) {
                break;
            }

            const colWidth = col.width;
            extendRight += colWidth;
            rightCols.push(currentCol);
            accumulatedWidth += colWidth;
            currentCol++;
        }
    } else if (horizontalAlign === 'right') {
        // 右对齐：向左延伸
        let currentCol = c - 1;
        let accumulatedWidth = w;

        while (accumulatedWidth < textWidth + 4 && currentCol >= 0) {
            const col = sheet.getCol(currentCol);

            // 跳过隐藏列
            if (col.hidden) {
                currentCol--;
                continue;
            }

            const key = `${r},${currentCol}`;
            const cellInfo = sheet.getCell(r, currentCol);

            // 如果相邻单元格非空、已被占用、或是合并单元格，停止延伸
            if (nonEmptyCells.has(key) || this.hiddenBorders.has(key) || cellInfo.isMerged) {
                break;
            }

            const colWidth = col.width;
            extendLeft += colWidth;
            leftCols.push(currentCol);
            accumulatedWidth += colWidth;
            currentCol--;
        }
    } else if (horizontalAlign === 'center') {
        // 居中对齐：双向延伸
        const neededWidth = textWidth + 4 - w;
        const halfNeeded = neededWidth / 2;

        // 向右延伸
        let currentCol = c + 1;
        let accumulatedWidth = 0;
        while (accumulatedWidth < halfNeeded && currentCol < sheet.colCount) {
            const col = sheet.getCol(currentCol);

            // 跳过隐藏列
            if (col.hidden) {
                currentCol++;
                continue;
            }

            const key = `${r},${currentCol}`;
            const cellInfo = sheet.getCell(r, currentCol);

            // 如果相邻单元格非空、已被占用、或是合并单元格，停止延伸
            if (nonEmptyCells.has(key) || this.hiddenBorders.has(key) || cellInfo.isMerged) {
                break;
            }

            const colWidth = col.width;
            extendRight += colWidth;
            rightCols.push(currentCol);
            accumulatedWidth += colWidth;
            currentCol++;
        }

        // 向左延伸
        currentCol = c - 1;
        accumulatedWidth = 0;
        while (accumulatedWidth < halfNeeded && currentCol >= 0) {
            const col = sheet.getCol(currentCol);

            // 跳过隐藏列
            if (col.hidden) {
                currentCol--;
                continue;
            }

            const key = `${r},${currentCol}`;
            const cellInfo = sheet.getCell(r, currentCol);

            // 如果相邻单元格非空、已被占用、或是合并单元格，停止延伸
            if (nonEmptyCells.has(key) || this.hiddenBorders.has(key) || cellInfo.isMerged) {
                break;
            }

            const colWidth = col.width;
            extendLeft += colWidth;
            leftCols.push(currentCol);
            accumulatedWidth += colWidth;
            currentCol--;
        }
    }

    // 记录被占用的单元格（包括主单元格和延伸的单元格）
    if (extendLeft > 0 || extendRight > 0) {
        // 根据延伸方向记录应该隐藏的边框
        if (extendRight > 0) {
            // 向右延伸：主单元格隐藏右边框
            const mainKey = `${r},${c}`;
            const existing = this.hiddenBorders.get(mainKey) || {};
            this.hiddenBorders.set(mainKey, { ...existing, right: true });

            // 右侧延伸单元格
            rightCols.forEach((col, index) => {
                const key = `${r},${col}`;
                const existing = this.hiddenBorders.get(key) || {};

                // 所有延伸单元格都隐藏左边框（进入边）
                existing.left = true;

                // 只有中间单元格隐藏右边框，最后一个不隐藏
                if (index < rightCols.length - 1) {
                    existing.right = true;
                }

                this.hiddenBorders.set(key, existing);
            });
        }

        if (extendLeft > 0) {
            // 向左延伸：主单元格隐藏左边框
            const mainKey = `${r},${c}`;
            const existing = this.hiddenBorders.get(mainKey) || {};
            this.hiddenBorders.set(mainKey, { ...existing, left: true });

            // 左侧延伸单元格
            leftCols.forEach((col, index) => {
                const key = `${r},${col}`;
                const existing = this.hiddenBorders.get(key) || {};

                // 所有延伸单元格都隐藏右边框（进入边）
                existing.right = true;

                // 只有中间单元格隐藏左边框，最后一个不隐藏
                if (index < leftCols.length - 1) {
                    existing.left = true;
                }

                this.hiddenBorders.set(key, existing);
            });
        }
    }

    // 使用扩展的裁剪区域绘制文本，但对齐基于原单元格
    const extendedX = x - extendLeft;
    const extendedW = w + extendLeft + extendRight;

    rText.call(this, {
        lines,
        cell,
        x: extendedX,
        y,
        w: extendedW,
        h,
        alignX: x,  // 对齐基于原单元格
        alignW: w,
        cfFormat,
        tableStyle,
        richText: !!richText
    });
}

// 渲染单元格（背景+文本）
export function rCell(x, y, w, h, cell) {
    rCellBackground.call(this, x, y, w, h, cell);
    rCellText.call(this, x, y, w, h, cell);
}

// ==================== 富文本辅助函数 ====================

// 按换行符拆分富文本 runs 为行数组，每行是 run 片段数组
function splitRichTextByNewline(runs) {
    const lines = [[]];
    for (const run of runs) {
        const parts = run.text.split('\n');
        for (let i = 0; i < parts.length; i++) {
            if (i > 0) lines.push([]);
            if (parts[i]) lines[lines.length - 1].push({ text: parts[i], font: run.font });
        }
    }
    return lines;
}

// 富文本自动换行
function wrapRichText(runs, maxWidth, cell) {
    const f = cell.font;
    const lines = [[]];
    let lineWidth = 0;

    for (const run of runs) {
        const rf = run.font || {};
        const bold = rf.bold ?? f.bold;
        const italic = rf.italic ?? f.italic;
        const size = ((rf.size ?? f.size ?? 11) * 1.333);
        const name = rf.name ?? f.name ?? '宋体';
        this.bc.font = `${bold ? 'bold' : ''} ${italic ? 'italic' : ''} ${size}px ${name}`;

        const chars = run.text.split('');
        let buf = '';
        for (const ch of chars) {
            if (ch === '\n') {
                if (buf) lines[lines.length - 1].push({ text: buf, font: run.font });
                buf = '';
                lines.push([]);
                lineWidth = 0;
                continue;
            }
            const cw = this.bc.measureText(ch).width;
            if (lineWidth + cw > maxWidth && lineWidth > 0) {
                if (buf) lines[lines.length - 1].push({ text: buf, font: run.font });
                buf = '';
                lines.push([]);
                lineWidth = 0;
            }
            buf += ch;
            lineWidth += cw;
        }
        if (buf) lines[lines.length - 1].push({ text: buf, font: run.font });
    }
    return lines;
}

// 测量富文本行总宽度
function measureRichLineWidth(richLine, cell) {
    const f = cell.font;
    let total = 0;
    for (const run of richLine) {
        const rf = run.font || {};
        const bold = rf.bold ?? f.bold;
        const italic = rf.italic ?? f.italic;
        const size = ((rf.size ?? f.size ?? 11) * 1.333);
        const name = rf.name ?? f.name ?? '宋体';
        this.bc.font = `${bold ? 'bold' : ''} ${italic ? 'italic' : ''} ${size}px ${name}`;
        total += this.bc.measureText(run.text).width;
    }
    return total;
}

// 富文本水平渲染
function rRichTextNormal(richLines, cell, baseX, y, baseW, h, defaultFsize, defaultColor) {
    const f = cell.font;
    const px = 2;
    const indent = cell.alignment?.indent || 0;
    const indentPx = indent * 9;
    const horizontalAlign = cell.horizontalAlign;
    const verticalAlign = cell.verticalAlign;

    // 计算每行的最大字号（px）
    const lineMaxSizes = richLines.map(runs => {
        let max = defaultFsize;
        for (const run of runs) {
            const s = ((run.font?.size ?? f.size ?? 11) * 1.333);
            if (s > max) max = s;
        }
        return max;
    });
    const lineHeightFactor = richLines.length > 1 ? 1.35 : 1.15;
    const lineHeights = lineMaxSizes.map(s => s * lineHeightFactor);
    const totalTextHeight = lineHeights.reduce((a, b) => a + b, 0);

    const savedBaseline = this.bc.textBaseline;
    this.bc.textBaseline = 'bottom';

    richLines.forEach((runs, lineIndex) => {
        if (!runs.length) return;

        const lineWidth = measureRichLineWidth.call(this, runs, cell);
        const maxSize = lineMaxSizes[lineIndex];

        // 计算行起始 X
        let startX;
        switch (horizontalAlign) {
            case 'center':
                startX = baseX + baseW / 2 - lineWidth / 2;
                break;
            case 'right':
                startX = baseX + (baseW - lineWidth) - px - 2 - indentPx;
                break;
            default:
                startX = baseX + px + indentPx;
                break;
        }

        // 计算行底部 Y（所有 run 底部对齐到此位置）
        const yBefore = lineHeights.slice(0, lineIndex).reduce((a, b) => a + b, 0);
        const yThisBottom = yBefore + lineHeights[lineIndex];
        let bottomY;
        switch (verticalAlign) {
            case 'center':
                bottomY = y + (h - totalTextHeight) / 2 + yThisBottom;
                break;
            case 'top':
                bottomY = y + px + yThisBottom;
                break;
            default:
                bottomY = y + h - totalTextHeight + yThisBottom - 1;
                break;
        }
        bottomY = Math.round(bottomY);

        // 逐 run 绘制
        let curX = Math.round(startX);
        for (const run of runs) {
            const rf = run.font || {};
            const bold = rf.bold ?? f.bold;
            const italic = rf.italic ?? f.italic;
            const size = ((rf.size ?? f.size ?? 11) * 1.333);
            const name = rf.name ?? f.name ?? '宋体';
            const color = rf.color ?? defaultColor;

            this.bc.font = `${bold ? 'bold' : ''} ${italic ? 'italic' : ''} ${size}px ${name}`;
            this.bc.fillStyle = color;
            this.bc.fillText(run.text, curX, bottomY);

            const runWidth = this.bc.measureText(run.text).width;

            // 删除线（文字垂直中间）
            if (rf.strike ?? f.strike) {
                this.bc.beginPath();
                this.bc.strokeStyle = color;
                this.bc.lineWidth = 1;
                const strikeY = bottomY - size * 0.35;
                this.bc.moveTo(curX, strikeY);
                this.bc.lineTo(curX + runWidth, strikeY);
                this.bc.stroke();
            }

            // 下划线（文字底部）
            const underline = rf.underline ?? f.underline;
            if (underline) {
                this.bc.beginPath();
                this.bc.strokeStyle = color;
                const ulY = bottomY + 1;
                if (underline === 'double') {
                    this.bc.lineWidth = 1;
                    this.bc.moveTo(curX, ulY - 2);
                    this.bc.lineTo(curX + runWidth, ulY - 2);
                    this.bc.stroke();
                    this.bc.beginPath();
                    this.bc.moveTo(curX, ulY);
                    this.bc.lineTo(curX + runWidth, ulY);
                } else {
                    this.bc.lineWidth = 1;
                    this.bc.moveTo(curX, ulY);
                    this.bc.lineTo(curX + runWidth, ulY);
                }
                this.bc.stroke();
            }

            curX += runWidth;
        }
    });

    this.bc.textBaseline = savedBaseline;
}

// 文本换行处理
export function wrapText(text, maxWidth, context) {
    const lines = [];
    const words = text.split('');
    let currentLine = words[0] || '';

    for (let i = 1; i < words.length; i++) {
        const word = words[i];
        const width = context.measureText(currentLine + word).width;
        if (width <= maxWidth) {
            currentLine += word;
        } else {
            lines.push(currentLine);
            currentLine = word;
        }
    }
    lines.push(currentLine);

    return lines;
}

// 调整文本位置
export function adjustTextPosition(horizontalAlign, verticalAlign, cellX, cellY, cellWidth, cellHeight, textWidth, textHeight, totalTextHeight, lineIndex, lineHeights, indent = 0) {
    const px = 2; // 单元格前后空格
    // 缩进：每级约3个字符宽度（@11pt字体约9px）
    const indentPx = indent * 9;
    let textX = cellX;
    let textY = cellY;

    // 处理水平对齐（缩进只对left/right生效）
    switch (horizontalAlign) {
        case 'center':
            textX = cellX + cellWidth / 2 - textWidth / 2;
            break;
        case 'right':
            textX = cellX + (cellWidth - textWidth) - px - 2 - indentPx;
            break;
        default:
            textX = cellX + px + indentPx;
            break;
    }

    // 处理垂直对齐
    switch (verticalAlign) {
        case 'center':
            textY = cellY + (cellHeight - totalTextHeight) / 2 + lineIndex * lineHeights + lineHeights - textHeight;
            break;
        case 'top':
            textY = cellY + lineIndex * lineHeights + px + lineHeights - textHeight;
            break;
        default:
            textY = cellY + cellHeight - totalTextHeight + lineIndex * lineHeights + lineHeights - textHeight - 1;
            break;
    }

    return { textX: Math.round(textX), textY: Math.round(textY) };
}

// 渲染文本
export function rText(info) {
    let { lines, cell, x, y, w, h, alignX, alignW, cfFormat, tableStyle, richText } = info;

    // 对齐基准区域（如果未指定，使用裁剪区域）
    const baseX = alignX !== undefined ? alignX : x;
    const baseW = alignW !== undefined ? alignW : w;

    const f = cell.font;
    let fsize = (f.size ?? 11) * 1.333;

    // 超级表表头加粗处理
    const tableBold = tableStyle?.isBold || false;
    const preferImportedTableStyle = tableStyle?._fromImport === true;
    const isBold = preferImportedTableStyle ? tableBold : (f.bold || tableBold);

    // 字体颜色优先级：条件格式 > 表格样式(当单元格为默认黑) > 单元格颜色 > 默认
    // 这样可保证导入的 Excel/WPS 表格样式颜色能覆盖基础 xf 字体色。
    const cfFontColor = cfFormat?.font?.color;
    const cellHasExplicitColor = cell._style?.font?.color;
    const tableFontColor = tableStyle?.fgColor;
    const cellFontColor = cell.font.color;
    const normalizedCellColor = typeof cellFontColor === 'string'
        ? (cellFontColor.length === 4
            ? `#${cellFontColor[1]}${cellFontColor[1]}${cellFontColor[2]}${cellFontColor[2]}${cellFontColor[3]}${cellFontColor[3]}`.toUpperCase()
            : cellFontColor.toUpperCase())
        : null;
    const canTableColorOverride = !!tableFontColor && (
        preferImportedTableStyle ||
        !cellHasExplicitColor ||
        normalizedCellColor === '#000000'
    );
    const textColor = cfFontColor ||
        (canTableColorOverride ? tableFontColor : null) ||
        (cellHasExplicitColor ? cellFontColor : null) ||
        tableFontColor ||
        cellFontColor ||
        '#333';
    const textRotation = cell.alignment?.textRotation;

    // 设置字体样式
    this.bc.font = `${isBold ? 'bold' : ''} ${f.italic ? 'italic' : ''} ${fsize}px ${f.name ?? '宋体'}`;

    // 富文本在竖排/旋转时 fallback 为纯文本
    if (richText && (textRotation === 255 || (textRotation && textRotation !== 0))) {
        lines = lines.map(line => Array.isArray(line) ? line.map(r => r.text).join('') : line);
        richText = false;
    }

    // shrinkToFit 处理：缩小字体以适应单元格宽度
    if (cell.alignment?.shrinkToFit && !cell.alignment?.wrapText && lines.length > 0) {
        const px = 4; // 左右边距
        let maxTextWidth;
        if (richText) {
            maxTextWidth = Math.max(...lines.map(line => measureRichLineWidth.call(this, line, cell)));
        } else {
            maxTextWidth = Math.max(...lines.map(line => this.bc.measureText(line).width));
        }
        const availableWidth = w - px;

        if (maxTextWidth > availableWidth && maxTextWidth > 0) {
            const scale = availableWidth / maxTextWidth;
            // 最小缩放到原字体的 50%
            const minScale = 0.5;
            const finalScale = Math.max(scale, minScale);
            fsize = fsize * finalScale;
            // 重新设置字体
            this.bc.font = `${isBold ? 'bold' : ''} ${f.italic ? 'italic' : ''} ${fsize}px ${f.name ?? '宋体'}`;
        }
    }

    // 应用裁剪区域
    this.bc.save();
    this.bc.beginPath();
    this.bc.rect(x, y, w, h);
    this.bc.clip();

    // 文字方向处理
    if (textRotation === 255) {
        // 竖排文字
        rTextVertical.call(this, lines, cell, x, y, w, h, fsize, textColor);
    } else if (textRotation && textRotation !== 0) {
        // 旋转文字
        rTextRotated.call(this, lines, cell, x, y, w, h, fsize, textColor, textRotation);
    } else if (richText) {
        // 富文本水平渲染
        rRichTextNormal.call(this, lines, cell, baseX, y, baseW, h, fsize, textColor);
    } else {
        // 普通水平文字
        rTextNormal.call(this, lines, cell, baseX, y, baseW, h, fsize, textColor);
    }

    this.bc.restore();
}

// 普通水平文字渲染
function rTextNormal(lines, cell, baseX, y, baseW, h, fsize, textColor) {
    const f = cell.font;
    const lineHeights = lines.length > 1 ? fsize * 1.35 : fsize * 1.15;
    const totalTextHeight = lines.length * lineHeights;
    const horizontalAlign = cell.horizontalAlign;
    const verticalAlign = cell.verticalAlign;
    const baseIndent = cell.alignment?.indent || 0;
    const pivotToggle = cell._pivotToggle;
    const hasPivotToggle = !!pivotToggle;
    const hasPivotIndent = Object.prototype.hasOwnProperty.call(cell, '_pivotIndent');
    const pivotIconSize = Math.min(12, Math.max(6, Math.floor(h - 4)));
    const iconSize = hasPivotToggle ? pivotIconSize : 0;
    const pivotIndentLevel = Number.isFinite(cell._pivotIndent) ? cell._pivotIndent : 0;
    const pivotIndentStep = pivotIndentLevel > 0 ? Math.max(10, Math.min(16, pivotIconSize + 4)) : 0;
    const pivotIndentPx = pivotIndentLevel > 0 ? pivotIndentLevel * pivotIndentStep : 0;
    const reservePivotIconSlot = hasPivotIndent || hasPivotToggle;
    const iconSlot = reservePivotIconSlot ? pivotIconSize + 4 : 0;
    const pivotFilter = cell._pivotFilter;
    const hasPivotFilter = !!pivotFilter;
    const indent = (hasPivotIndent || hasPivotToggle || hasPivotFilter) ? 0 : baseIndent;
    const filterSize = hasPivotFilter ? Math.min(12, Math.max(6, Math.floor(h - 4))) : 0;
    const filterSlot = hasPivotFilter ? filterSize + 6 : 0;
    let pivotDrawn = false;
    let filterDrawn = false;

    if (!hasPivotToggle && cell._pivotToggleRect) {
        delete cell._pivotToggleRect;
    }
    if (hasPivotToggle && iconSize <= 0 && cell._pivotToggleRect) {
        delete cell._pivotToggleRect;
    }
    if (!hasPivotFilter && cell._pivotFilterRect) {
        delete cell._pivotFilterRect;
    }
    if (hasPivotFilter && filterSize <= 0 && cell._pivotFilterRect) {
        delete cell._pivotFilterRect;
    }

    // 检查是否是会计格式
    const accounting = cell.accountingData;
    if (accounting && lines.length === 1) {
        rTextAccounting.call(this, accounting, cell, baseX, y, baseW, h, fsize, textColor, lineHeights, totalTextHeight, verticalAlign);
        return;
    }

    lines.forEach((line, index) => {
        const textOffset = pivotIndentPx + iconSlot;
        const isLeftAlign = !horizontalAlign || horizontalAlign === 'left';
        const textBaseX = isLeftAlign ? baseX + textOffset : baseX;
        const textBaseW = isLeftAlign ? Math.max(0, baseW - textOffset - filterSlot) : Math.max(0, baseW - filterSlot);
        const textWidth = this.bc.measureText(line).width;
        const { textX, textY } = adjustTextPosition.call(this, horizontalAlign, verticalAlign, textBaseX, y, textBaseW, h, textWidth, fsize, totalTextHeight, index, lineHeights, indent);

        this.bc.fillStyle = textColor;
        if (hasPivotToggle && !pivotDrawn && iconSize > 0) {
            let iconY = y + (h - iconSize) / 2;
            if (verticalAlign === 'top') {
                iconY = y + 2;
            } else if (verticalAlign === 'bottom') {
                iconY = y + h - iconSize - 2;
            }
            const iconX = baseX + 2 + pivotIndentPx;
            drawPivotToggleIcon(this.bc, iconX, iconY, iconSize, !!pivotToggle?.collapsed);
            cell._pivotToggleRect = { x: iconX, y: iconY, w: iconSize, h: iconSize };
            pivotDrawn = true;
        }
        if (hasPivotFilter && !filterDrawn && filterSize > 0) {
            let iconY = y + (h - filterSize) / 2;
            if (verticalAlign === 'top') {
                iconY = y + 2;
            } else if (verticalAlign === 'bottom') {
                iconY = y + h - filterSize - 2;
            }
            const iconX = baseX + baseW - filterSize - 2;
            this.drawFilterIcon(this.bc, iconX, iconY, filterSize, !!pivotFilter?.active, null);
            cell._pivotFilterRect = { x: iconX, y: iconY, w: filterSize, h: filterSize };
            filterDrawn = true;
        }
        this.bc.fillText(line, textX, textY);

        // 删除线
        if (f.strike && line) {
            this.bc.beginPath();
            this.bc.strokeStyle = textColor;
            this.bc.lineWidth = 1;
            const strikeY = textY + fsize * 0.4;
            this.bc.moveTo(textX, strikeY);
            this.bc.lineTo(textX + textWidth, strikeY);
            this.bc.stroke();
        }

        // 下划线
        if (f.underline && line) {
            this.bc.beginPath();
            this.bc.strokeStyle = textColor;
            const underlineY = textY + fsize;
            if (f.underline === "single") {
                this.bc.lineWidth = 1;
                this.bc.moveTo(textX, underlineY);
                this.bc.lineTo(textX + textWidth, underlineY);
            } else if (f.underline === "double") {
                this.bc.lineWidth = 1;
                this.bc.moveTo(textX, underlineY - 2);
                this.bc.lineTo(textX + textWidth, underlineY - 2);
                this.bc.stroke();
                this.bc.beginPath();
                this.bc.moveTo(textX, underlineY);
                this.bc.lineTo(textX + textWidth, underlineY);
            }
            this.bc.stroke();
        }
    });
}

// 会计格式渲染：货币符号左对齐，数字右对齐
function rTextAccounting(accounting, cell, baseX, y, baseW, h, fsize, textColor, lineHeights, totalTextHeight, verticalAlign) {
    const px = 2; // 单元格内边距
    const { prefix, number, suffix } = accounting;

    // 计算各部分宽度
    const prefixWidth = prefix ? this.bc.measureText(prefix).width : 0;
    const numberText = number + suffix;
    const numberWidth = this.bc.measureText(numberText).width;
    const totalWidth = prefixWidth + numberWidth + px * 2;

    // 计算垂直位置（与 adjustTextPosition 保持一致）
    let textY;
    switch (verticalAlign) {
        case 'center':
            textY = y + (h - totalTextHeight) / 2 + lineHeights - fsize;
            break;
        case 'top':
            textY = y + px + lineHeights - fsize;
            break;
        default: // bottom
            textY = y + h - totalTextHeight + lineHeights - fsize - 1;
            break;
    }

    this.bc.fillStyle = textColor;

    // 如果宽度不够，显示 ### 占位
    if (totalWidth > baseW) {
        const hashText = '#'.repeat(Math.max(1, Math.floor((baseW - px * 2) / this.bc.measureText('#').width)));
        const hashWidth = this.bc.measureText(hashText).width;
        const hashX = baseX + baseW - hashWidth - px;
        this.bc.fillText(hashText, hashX, textY);
        return;
    }

    // 绘制前缀（货币符号）- 左对齐
    if (prefix) {
        const prefixX = baseX + px;
        this.bc.fillText(prefix, prefixX, textY);
    }

    // 绘制数字 + 后缀 - 右对齐
    const numberX = baseX + baseW - numberWidth - px;
    this.bc.fillText(numberText, numberX, textY);
}

// 竖排文字渲染
function rTextVertical(lines, cell, x, y, w, h, fsize, textColor) {
    const text = lines.join('');
    const chars = [...text]; // 支持emoji等多字节字符
    const charHeight = fsize * 1.1; // 字符间距
    const totalHeight = chars.length * charHeight;
    const horizontalAlign = cell.horizontalAlign;
    const verticalAlign = cell.verticalAlign;

    // 计算起始X位置
    let startX;
    switch (horizontalAlign) {
        case 'left': startX = x + fsize / 2 + 2; break;
        case 'right': startX = x + w - fsize / 2 - 2; break;
        default: startX = x + w / 2; break;
    }

    // 计算起始Y位置（使用top基线，y即字符顶部）
    let startY;
    switch (verticalAlign) {
        case 'top': startY = y + 2; break;
        case 'center': startY = y + (h - totalHeight) / 2; break;
        default: startY = y + h - totalHeight - 2; break;
    }

    this.bc.fillStyle = textColor;
    this.bc.textAlign = 'center';
    this.bc.textBaseline = 'top';

    chars.forEach((char, i) => {
        this.bc.fillText(char, startX, startY + i * charHeight);
    });

    this.bc.textAlign = 'left';
    this.bc.textBaseline = 'alphabetic';
}

// 旋转文字渲染
function rTextRotated(lines, cell, x, y, w, h, fsize, textColor, rotation) {
    const text = lines.join('');
    const textWidth = this.bc.measureText(text).width;
    const horizontalAlign = cell.horizontalAlign;
    const verticalAlign = cell.verticalAlign;

    // 计算旋转角度（弧度）
    // Excel: 1-90 逆时针，91-180 顺时针（值-90度）
    let angle;
    if (rotation <= 90) {
        angle = -rotation * Math.PI / 180; // 逆时针为负
    } else {
        angle = (rotation - 90) * Math.PI / 180; // 顺时针为正
    }

    // 计算旋转后文字的边界框
    const cos = Math.abs(Math.cos(angle));
    const sin = Math.abs(Math.sin(angle));
    const rotatedW = textWidth * cos + fsize * sin;
    const rotatedH = textWidth * sin + fsize * cos;

    // 计算中心点位置
    let cx, cy;
    switch (horizontalAlign) {
        case 'left': cx = x + rotatedW / 2 + 2; break;
        case 'right': cx = x + w - rotatedW / 2 - 2; break;
        default: cx = x + w / 2; break;
    }
    switch (verticalAlign) {
        case 'top': cy = y + rotatedH / 2 + 2; break;
        case 'center': cy = y + h / 2; break;
        default: cy = y + h - rotatedH / 2 - 2; break;
    }

    this.bc.save();
    this.bc.translate(cx, cy);
    this.bc.rotate(angle);
    this.bc.fillStyle = textColor;
    this.bc.textAlign = 'center';
    this.bc.textBaseline = 'middle';
    this.bc.fillText(text, 0, 0);
    this.bc.restore();
}

// 渲染合并的单元格
export function rMerged(sheet) {
    // 遍历所有合并单元格
    sheet.merges.forEach(m => {
        const { s: start, e: end } = m;
        const sCell = sheet.getAreaInviewInfo(m);
        const cellInfo = sheet.getCell(start.r, start.c);
        if (sCell.inView) { // 区域有一部分在视图中,那么合并
            rCell.call(this, sCell.x, sCell.y, sCell.w, sCell.h, cellInfo);
            // ****************设置合并单元格边框比较特殊,逻辑为：比如如果top有边框,那么top的所有单元格的topborder属性都必须一样****************
            // 获取合并区域的范围
            const { s, e } = m; // s是起始位置,e是结束位置
            const startRow = s.r;
            const startCol = s.c;
            const endRow = e.r;
            const endCol = e.c;
            // 存储四边的边框属性
            const result = {};
            // 检查顶部边框
            let topBorderConsistent = true;
            const topBorder = sheet.getCell(startRow, startCol).border.top;
            for (let c = startCol + 1; c <= endCol; c++) {
                const cellBorder = sheet.getCell(startRow, c).border.top;
                if (!cellBorder || cellBorder.color !== topBorder.color || cellBorder.style !== topBorder.style) {
                    topBorderConsistent = false;
                    break;
                }
            }
            if (topBorderConsistent && topBorder) {
                result.top = {
                    color: topBorder.color,
                    style: topBorder.style
                };
            }
            // 检查右侧边框
            let rightBorderConsistent = true;
            const rightBorder = sheet.getCell(startRow, endCol).border.right;
            for (let r = startRow + 1; r <= endRow; r++) {
                const cellBorder = sheet.getCell(r, endCol).border.right;
                if (!cellBorder || cellBorder.color !== rightBorder.color || cellBorder.style !== rightBorder.style) {
                    rightBorderConsistent = false;
                    break;
                }
            }
            if (rightBorderConsistent && rightBorder) {
                result.right = {
                    color: rightBorder.color,
                    style: rightBorder.style
                };
            }
            // 检查底部边框
            let bottomBorderConsistent = true;
            const bottomBorder = sheet.getCell(endRow, startCol).border.bottom;
            for (let c = startCol + 1; c <= endCol; c++) {
                const cellBorder = sheet.getCell(endRow, c).border.bottom;
                if (!cellBorder || cellBorder.color !== bottomBorder.color || cellBorder.style !== bottomBorder.style) {
                    bottomBorderConsistent = false;
                    break;
                }
            }
            if (bottomBorderConsistent && bottomBorder) {
                result.bottom = {
                    color: bottomBorder.color,
                    style: bottomBorder.style
                };
            }
            // 检查左侧边框
            let leftBorderConsistent = true;
            const leftBorder = sheet.getCell(startRow, startCol).border.left;
            for (let r = startRow + 1; r <= endRow; r++) {
                const cellBorder = sheet.getCell(r, startCol).border.left;
                if (!cellBorder || cellBorder.color !== leftBorder.color || cellBorder.style !== leftBorder.style) {
                    leftBorderConsistent = false;
                    break;
                }
            }
            if (leftBorderConsistent && leftBorder) {
                result.left = {
                    color: leftBorder.color,
                    style: leftBorder.style
                };
            }
            const masterBorder = sheet.getCell(startRow, startCol).border;
            if (masterBorder.diagonal) {
                result.diagonal = {
                    color: masterBorder.diagonal.color,
                    style: masterBorder.diagonal.style
                };
                if (masterBorder.diagonalDown) result.diagonalDown = true;
                if (masterBorder.diagonalUp) result.diagonalUp = true;
            }
            // *******四周边框检测结束*******

            if (Object.keys(result).length > 0) {
                this.borders.push({
                    b: result,
                    x: sCell.x,
                    y: sCell.y,
                    w: sCell.w,
                    h: sCell.h,
                    r: startRow,
                    c: startCol
                }); // 边框待渲染
            }
        }
    });
}

// 边框渲染
export function rBorder(ciArr) {
    const snap = (value) => this.snap(value);
    const px = (value) => this.px(value);

    // 定义边框样式和对应的线宽
    const borderStyles = {
        thin: { lineWidth: 1 },
        dotted: { lineWidth: 1, lineDash: [1, 2] },
        dashDot: { lineWidth: 1, lineDash: [4, 2, 1, 2] },
        hair: { lineWidth: 0.5 },
        dashDotDot: { lineWidth: 1, lineDash: [4, 2, 1, 2, 1, 2] },
        slantDashDot: { lineWidth: 1, lineDash: [5, 5] }, // 示例,实际应用中请调整为斜线点划样式
        mediumDashed: { lineWidth: 1.5, lineDash: [7, 5] },
        mediumDashDotDot: { lineWidth: 1.5, lineDash: [4, 2, 1, 2, 1, 2] },
        mediumDashDot: { lineWidth: 1.5, lineDash: [4, 2, 1, 2] },
        medium: { lineWidth: 1.5 },
        double: { lineWidth: 1 },
        thick: { lineWidth: 2.5 },
        dashed: { lineWidth: 1, lineDash: [5, 5] } // 这里添加了新的样式
    };

    // 边框绘制函数
    const drawBorder = (bc, x, y, w, h, position, style, color) => {
        const styleInfo = borderStyles[style]
        if (!styleInfo) return; // 没有样式信息时退出
        const x1 = snap(x);
        const y1 = snap(y);
        const x2 = snap(x + w);
        const y2 = snap(y + h);
        const offset = px(1);
        const isHairline = style === 'hair';
        bc.save(); // 保存当前的绘图状态
        bc.beginPath();
        bc.strokeStyle = color;
        bc.lineWidth = isHairline ? px(1) : px(styleInfo.lineWidth);
        if (isHairline) {
            bc.globalAlpha = 0.5;
        }
        if (styleInfo.lineDash) {
            bc.setLineDash(styleInfo.lineDash.map(px));
        } else {
            bc.setLineDash([]); // 清除虚线设置
        }
        const drawLine = (x1, y1, x2, y2) => {
            bc.moveTo(x1, y1);
            bc.lineTo(x2, y2);
        };
        const drawDoubleDiagonal = (xStart, yStart, xEnd, yEnd) => {
            const dx = xEnd - xStart;
            const dy = yEnd - yStart;
            const len = Math.hypot(dx, dy) || 1;
            const ux = (dy / len) * offset;
            const uy = (-dx / len) * offset;
            drawLine(snap(xStart - ux), snap(yStart - uy), snap(xEnd - ux), snap(yEnd - uy));
            drawLine(snap(xStart + ux), snap(yStart + uy), snap(xEnd + ux), snap(yEnd + uy));
        };
        switch (position) {
            case 'top':
                if (style === 'double') {
                    drawLine(x1, snap(y1 - offset), x2, snap(y1 - offset));
                    drawLine(x1, snap(y1 + offset), x2, snap(y1 + offset));
                } else {
                    drawLine(x1, y1, x2, y1);
                }
                break;
            case 'bottom':
                if (style === 'double') {
                    drawLine(x1, snap(y2 - offset), x2, snap(y2 - offset));
                    drawLine(x1, snap(y2 + offset), x2, snap(y2 + offset));
                } else {
                    drawLine(x1, y2, x2, y2);
                }
                break;
            case 'left':
                if (style === 'double') {
                    drawLine(snap(x1 - offset), y1, snap(x1 - offset), y2);
                    drawLine(snap(x1 + offset), y1, snap(x1 + offset), y2);
                } else {
                    drawLine(x1, y1, x1, y2);
                }
                break;
            case 'right':
                if (style === 'double') {
                    drawLine(snap(x2 - offset), y1, snap(x2 - offset), y2);
                    drawLine(snap(x2 + offset), y1, snap(x2 + offset), y2);
                } else {
                    drawLine(x2, y1, x2, y2);
                }
                break;
            case 'diagonalDown':
                if (style === 'double') {
                    drawDoubleDiagonal(x1, y1, x2, y2);
                } else {
                    drawLine(x1, y1, x2, y2);
                }
                break;
            case 'diagonalUp':
                if (style === 'double') {
                    drawDoubleDiagonal(x1, y2, x2, y1);
                } else {
                    drawLine(x1, y2, x2, y1);
                }
                break;
        }
        bc.stroke();
        bc.closePath();
        bc.restore(); // 恢复之前保存的绘图状态
    };

    // 应用边框样式
    // { b: cellInfo.border, x: _x, y: _y, w: nowCellWidth, h: rowHeight, r: r, c: c }
    ciArr.forEach(info => {
        // 获取该单元格需要隐藏的边框信息
        const hiddenInfo = this.hiddenBorders && info.r !== undefined && info.c !== undefined
            ? this.hiddenBorders.get(`${info.r},${info.c}`)
            : null;

        ['top', 'bottom', 'left', 'right'].forEach(position => {
            const border = info.b[position];

            // 检查该边是否需要隐藏
            if (hiddenInfo && hiddenInfo[position]) {
                return; // 隐藏该边框
            }

            if (borderStyles[border?.style]) {
                drawBorder(this.bc, info.x, info.y, info.w, info.h, position, border.style, border.color);
            }
        });
        const diagonal = info.b.diagonal;
        if (borderStyles[diagonal?.style]) {
            const drawDiagonalUp = info.b.diagonalUp === true;
            const drawDiagonalDown = info.b.diagonalDown === true || (!drawDiagonalUp && info.b.diagonalDown === undefined);
            if (drawDiagonalDown) {
                drawBorder(this.bc, info.x, info.y, info.w, info.h, 'diagonalDown', diagonal.style, diagonal.color);
            }
            if (drawDiagonalUp) {
                drawBorder(this.bc, info.x, info.y, info.w, info.h, 'diagonalUp', diagonal.style, diagonal.color);
            }
        }
    });

}

// 渲染基础数据
export function rBaseData(sheet) {
    this.bc.strokeStyle = '#ddd';
    this.bc.textAlign = 'left';
    this.bc.textBaseline = 'top';

    this.borders = []; // 准备渲染的边框
    this.hiddenBorders = new Map(); // key: "r,c", value: { left: bool, right: bool } 记录具体隐藏哪条边

    // 用于存储单元格位置信息和非空单元格
    const cellPositions = new Map(); // key: "r,c", value: { x, y, w, h, cellInfo, row, col, cfFormat }
    const nonEmptyCells = new Set(); // key: "r,c"

    // 条件格式预计算
    const cf = sheet.CF;
    const hasCF = cf && cf.rules.length > 0;
    let rangeDataMap = null;
    if (hasCF) {
        // 预计算需要范围数据的规则
        rangeDataMap = new Map();
        cf.rules.forEach(rule => {
            if (rule.needsRangeData && !rangeDataMap.has(rule.sqref)) {
                rangeDataMap.set(rule.sqref, rule.getRangeData());
            }
        });
    }

    // 第一遍：绘制背景，收集单元格信息
    let y = sheet.headHeight;
    sheet.vi.rowsIndex.forEach(r => {
        let x = sheet.indexWidth;
        let rowHeight = sheet.getRow(r).height;
        const row = sheet.getRow(r);

        sheet.vi.colsIndex.forEach(c => {
            let nowCellWidth = sheet.getCol(c).width;
            const cellInfo = sheet.getCell(r, c);

            if (!cellInfo.isMerged) {
                const _x = x;
                const _y = y;

                // 获取条件格式结果
                const cfFormat = hasCF ? cf.getFormat(r, c) : null;

                // 获取超级表样式
                const tableStyle = getTableCellStyle(sheet, r, c);

                // 绘制背景（合并条件格式背景和超级表样式）
                rCellBackground.call(this, _x, _y, nowCellWidth, rowHeight, cellInfo, true, cfFormat, tableStyle);

                // 收集单元格位置信息（包含条件格式结果和超级表样式）
                cellPositions.set(`${r},${c}`, {
                    x: _x,
                    y: _y,
                    w: nowCellWidth,
                    h: rowHeight,
                    cellInfo,
                    row: r,
                    col: c,
                    cfFormat,
                    tableStyle
                });

                // 收集非空单元格
                if (cellInfo.showVal && cellInfo.showVal.toString()) {
                    nonEmptyCells.add(`${r},${c}`);
                }

                // 收集边框信息（叠加条件格式边框）
                const tableBorderColor = tableStyle?._fromImport === true ? tableStyle?.borderColor : null;
                const baseBorder = tableBorderColor
                    ? applyTableBorderColor(cellInfo.border || {}, tableBorderColor)
                    : (cellInfo.border || {});
                const mergedBorder = cfFormat?.border ? { ...baseBorder, ...cfFormat.border } : baseBorder;
                if (Object.keys(mergedBorder).length > 0) {
                    this.borders.push({
                        b: mergedBorder,
                        x: _x,
                        y: _y,
                        w: nowCellWidth,
                        h: rowHeight,
                        r: r,
                        c: c
                    });
                }
            }

            x += nowCellWidth;
        });

        y += rowHeight;
    });

    // 存储 cellPositions 供 rConditionalFormats 使用（图标）
    this._cfCellPositions = cellPositions;

    // 第二遍：绘制数据条（在文字之前，避免覆盖）
    cellPositions.forEach((pos, key) => {
        const { x, y, w, h, cfFormat } = pos;
        if (cfFormat?.dataBar) {
            this.rDataBar(x, y, w, h, cfFormat.dataBar);
        }
    });

    // 第三遍：绘制文本（带溢出支持，合并条件格式字体色和超级表样式）
    cellPositions.forEach((pos, key) => {
        const { x, y, w, h, cellInfo, row, col, cfFormat, tableStyle } = pos;

        // 只渲染非空且非合并的单元格文本
        if (cellInfo.showVal && cellInfo.showVal.toString() && !cellInfo.isMerged) {
            rCellTextWithOverflow.call(this, row, col, x, y, w, h, cellInfo, sheet, nonEmptyCells, cfFormat, tableStyle);
        }
    });

    // 第三遍：绘制网格线（被占用的单元格只隐藏相应的边框）
    if (sheet.showGridLines) {
        const lineWidth = this.px(1);
        cellPositions.forEach((pos, key) => {
            const { x, y, w, h, cellInfo, cfFormat } = pos;
            const hiddenInfo = this.hiddenBorders.get(key);
            const x1 = this.snap(x);
            const y1 = this.snap(y);
            const x2 = this.snap(x + w);
            const y2 = this.snap(y + h);

            // 网格线颜色：条件格式背景色 > 单元格背景色 > 默认灰色
            const cfFillColor = cfFormat?.fill?.fgColor;
            this.bc.strokeStyle = cfFillColor || cellInfo.fill.fgColor || '#ddd';
            this.bc.lineWidth = lineWidth;
            this.bc.beginPath();

            if (!hiddenInfo) {
                // 没有隐藏信息，绘制完整网格线
                this.bc.strokeRect(x1, y1, x2 - x1, y2 - y1);
            } else {
                // 有隐藏信息，绘制非隐藏的边
                if (!hiddenInfo.top) {
                    this.bc.moveTo(x1, y1);
                    this.bc.lineTo(x2, y1);
                }
                if (!hiddenInfo.bottom) {
                    this.bc.moveTo(x1, y2);
                    this.bc.lineTo(x2, y2);
                }
                if (!hiddenInfo.left) {
                    this.bc.moveTo(x1, y1);
                    this.bc.lineTo(x1, y2);
                }
                if (!hiddenInfo.right) {
                    this.bc.moveTo(x2, y1);
                    this.bc.lineTo(x2, y2);
                }
                this.bc.stroke();
            }
        });
    }

    // 记录可视区域是否有内容（供 AI 等模块使用）
    sheet.vi.hasContent = nonEmptyCells.size > 0;
    
}

function drawPivotToggleIcon(ctx, x, y, size, collapsed) {
    if (size <= 0) return;
    const rectX = Math.floor(x) + 0.5;
    const rectY = Math.floor(y) + 0.5;
    ctx.save();
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(rectX, rectY, size, size);
    ctx.strokeStyle = '#8a8a8a';
    ctx.lineWidth = 1;
    ctx.strokeRect(rectX, rectY, size, size);
    const cx = rectX + size / 2;
    const cy = rectY + size / 2;
    const half = Math.max(2, Math.floor(size * 0.25));
    ctx.strokeStyle = '#4a4a4a';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(cx - half, cy);
    ctx.lineTo(cx + half, cy);
    if (collapsed) {
        ctx.moveTo(cx, cy - half);
        ctx.lineTo(cx, cy + half);
    }
    ctx.stroke();
    ctx.restore();
}



