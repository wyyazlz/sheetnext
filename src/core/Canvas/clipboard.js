/**
 * 剪贴板工具类 - 高性能样式转换与缓存
 * 用于Excel风格的复制粘贴（HTML+文本双格式）
 */

// 样式缓存，避免重复计算
const styleCache = new Map();
const colorCache = new Map();

// 缓存大小限制（防止内存泄漏）
const MAX_CACHE_SIZE = 1000;

// 自动清理缓存
function checkCacheSize(cache) {
    if (cache.size > MAX_CACHE_SIZE) {
        // 清除最旧的一半缓存
        const keysToDelete = Array.from(cache.keys()).slice(0, Math.floor(MAX_CACHE_SIZE / 2));
        keysToDelete.forEach(key => cache.delete(key));
    }
}

/**
 * 将十六进制颜色转换为RGB或保持原格式
 */
function normalizeColor(color) {
    if (!color) return '';
    if (colorCache.has(color)) return colorCache.get(color);

    let result = color;
    // 处理#AARRGGBB格式（带透明度）
    if (color.length === 9 && color.startsWith('#')) {
        const alpha = parseInt(color.slice(1, 3), 16) / 255;
        const r = parseInt(color.slice(3, 5), 16);
        const g = parseInt(color.slice(5, 7), 16);
        const b = parseInt(color.slice(7, 9), 16);
        result = `rgba(${r},${g},${b},${alpha.toFixed(2)})`;
    }
    // #RRGGBB格式保持不变

    colorCache.set(color, result);
    return result;
}

/**
 * 将Cell样式转换为CSS类样式（Excel兼容格式）
 */
export function cellStyleToCSS(cell) {
    try {
        // 使用缓存键
        const cacheKey = JSON.stringify({
            font: cell.font,
            fill: cell.fill,
            border: cell.border,
            alignment: cell.alignment
        });

        if (styleCache.has(cacheKey)) {
            return styleCache.get(cacheKey);
        }

        checkCacheSize(styleCache); // 检查缓存大小

        const styles = [];

        // 背景填充
        if (cell.fill?.fgColor) {
            styles.push(`background-color:${normalizeColor(cell.fill.fgColor)}`);
        }

        // 字体颜色
        if (cell.font?.color) {
            styles.push(`color:${normalizeColor(cell.font.color)}`);
        }

        // 字体样式（使用简写形式：font: style weight size family）
        if (cell.font) {
            const f = cell.font;
            const fontStyle = f.italic ? 'italic' : 'normal';
            const fontWeight = f.bold ? 'bold' : 'normal';
            const fontSize = f.size ? `${Number(f.size).toFixed(2)}pt` : '11.00pt';
            const fontFamily = f.name || 'Arial';
            styles.push(`font:${fontStyle} ${fontWeight} ${fontSize} ${fontFamily}`);

            // 下划线和删除线单独处理
            if (f.underline || f.strike) {
                const decorations = [];
                if (f.underline) decorations.push('underline');
                if (f.strike) decorations.push('line-through');
                styles.push(`text-decoration:${decorations.join(' ')}`);
            }
        }

        // 对齐方式
        const wrapText = cell.alignment?.wrapText;
        styles.push(`white-space:${wrapText ? 'normal' : 'nowrap'}`);

        if (cell.alignment?.vertical) {
            const valign = cell.alignment.vertical === 'center' ? 'middle' :
                          cell.alignment.vertical === 'bottom' ? 'bottom' : 'top';
            styles.push(`vertical-align:${valign}`);
        }

        if (cell.alignment?.horizontal) {
            const align = cell.alignment.horizontal === 'center' ? 'center' :
                         cell.alignment.horizontal === 'right' ? 'right' : 'left';
            styles.push(`text-align:${align}`);
        }

        // 边框（使用pt单位，Excel兼容）
        if (cell.border) {
            const b = cell.border;
            ['top', 'right', 'bottom', 'left'].forEach(side => {
                if (b[side]) {
                    const style = b[side].style || 'thin';
                    const color = normalizeColor(b[side].color) || '#000';
                    // Excel使用pt单位：thin=0.75pt, medium=1.5pt, thick=2.25pt
                    const width = style === 'medium' ? '1.5pt' : style === 'thick' ? '2.25pt' : '0.75pt';
                    styles.push(`border-${side}:${width} solid ${color}`);
                }
            });
        }

        // Microsoft Office 特定属性
        styles.push('mso-number-format:General');

        const result = styles.join(';\n');
        styleCache.set(cacheKey, result);
        return result;
    } catch (error) {
        console.error('样式转换失败:', error);
        return '';
    }
}

/**
 * 生成HTML表格（使用CSS类样式，Excel兼容格式）
 */
export function generateClipboardHTML(sheet, area) {
    try {
        const styleMap = new Map(); // 样式 -> 类名映射
        const cellDataMap = new Map(); // 存储每个单元格的数据和类名
        let styleCounter = 1;

        // 第一遍：收集所有单元格的样式并生成类名
        for (let r = area.s.r; r <= area.e.r; r++) {
            for (let c = area.s.c; c <= area.e.c; c++) {
                const cell = sheet.getCell(r, c);

                // 跳过已合并的子单元格（只保留主单元格）
                // 主单元格的特征：cell.master 指向自己
                if (cell.isMerged && cell.master && !(cell.master.r === r && cell.master.c === c)) {
                    continue;
                }

                const styleCSS = cellStyleToCSS(cell);

                // 如果这个样式是新的，分配一个类名
                if (!styleMap.has(styleCSS)) {
                    styleMap.set(styleCSS, `sjs${styleCounter++}`);
                }

                cellDataMap.set(`${r}-${c}`, {
                    value: escapeHtml(cell.showVal ?? ''),
                    className: styleMap.get(styleCSS),
                    cell: cell
                });
            }
        }

        // 生成 CSS 样式表
        const cssRules = [];
        styleMap.forEach((className, styleCSS) => {
            cssRules.push(`.${className}{\n${styleCSS}\n}`);
        });

        // 第二遍：生成表格行
        const rows = [];
        for (let r = area.s.r; r <= area.e.r; r++) {
            const cells = [];

            for (let c = area.s.c; c <= area.e.c; c++) {
                const cellData = cellDataMap.get(`${r}-${c}`);
                if (!cellData) continue; // 跳过合并的子单元格

                const cell = cellData.cell;

                // 处理合并单元格
                let colspan = 1, rowspan = 1;
                // 主单元格的特征：isMerged 且 master 指向自己
                if (cell.isMerged && cell.master && cell.master.r === r && cell.master.c === c) {
                    const mergeInfo = sheet.merges.find(m =>
                        m.s.r === r && m.s.c === c
                    );
                    if (mergeInfo) {
                        colspan = mergeInfo.e.c - mergeInfo.s.c + 1;
                        rowspan = mergeInfo.e.r - mergeInfo.s.r + 1;
                    }
                }

                const colspanAttr = colspan > 1 ? ` colSpan=${colspan}` : '';
                const rowspanAttr = rowspan > 1 ? ` rowSpan=${rowspan}` : '';

                cells.push(
                    `<td${rowspanAttr}${colspanAttr} class="${cellData.className}">${cellData.value}</td>`
                );
            }

            if (cells.length > 0) {
                rows.push(`<tr>${cells.join('')}</tr>`);
            } else {
                rows.push('<tr></tr>'); // 空行（合并单元格跨越的行）
            }
        }

        // 生成完整的 HTML（Excel 兼容格式）
        return `<html>
<head>
<style>
${cssRules.join('\n')}
</style>
</head>
<body>
<table gc-sjs-clipboard="true" style="min-width:1px;min-height:1px;">${rows.join('')}</table>
</body>
</html>`;
    } catch (error) {
        console.error('生成HTML失败:', error);
        // 返回空table而非抛出错误
        return '<table></table>';
    }
}

/**
 * 生成纯文本（Tab分隔）
 */
export function generateClipboardText(sheet, area) {
    try {
        const rows = [];

        for (let r = area.s.r; r <= area.e.r; r++) {
            const cells = [];
            for (let c = area.s.c; c <= area.e.c; c++) {
                cells.push(sheet.getCell(r, c).editVal ?? '');
            }
            rows.push(cells.join('\t'));
        }

        return rows.join('\r\n');
    } catch (error) {
        console.error('生成纯文本失败:', error);
        return '';
    }
}

/**
 * 从HTML解析样式（支持CSS类样式）
 */
const parseCache = new Map();

export function parseHTMLStyle(htmlString) {
    try {
        if (parseCache.has(htmlString)) {
            return parseCache.get(htmlString);
        }

        checkCacheSize(parseCache); // 检查缓存大小

        // 创建临时容器并插入到真实DOM中（隐藏），以便CSS样式被正确应用
        const tempContainer = document.createElement('div');
        tempContainer.style.cssText = 'position:absolute;left:-9999px;top:-9999px;visibility:hidden;';
        document.body.appendChild(tempContainer);

        // 将HTML插入到容器中，让浏览器应用CSS样式
        tempContainer.innerHTML = htmlString;
        const table = tempContainer.querySelector('table');

        if (!table) {
            document.body.removeChild(tempContainer);
            parseCache.set(htmlString, null);
            return null;
        }

        const result = {
            data: [],
            styles: [],
            merges: [] // 新增：存储合并单元格信息
        };

        const rows = table.querySelectorAll('tr');
        let currentRowIndex = 0; // 实际数据行索引

        rows.forEach((tr) => {
            const cells = tr.querySelectorAll('td, th');

            // 跳过空行
            if (cells.length === 0) {
                return;
            }

            const rowData = [];
            const rowStyles = [];
            let currentColIndex = 0; // 实际数据列索引

            cells.forEach((td) => {
                // 获取单元格值
                const value = td.textContent || '';

                // 解析样式（使用计算后的样式）
                const style = parseCellStyle(td);

                // 处理合并单元格
                const rowspan = parseInt(td.getAttribute('rowspan') || td.getAttribute('rowSpan')) || 1;
                const colspan = parseInt(td.getAttribute('colspan') || td.getAttribute('colSpan')) || 1;

                // 如果是合并单元格，记录合并信息
                if (rowspan > 1 || colspan > 1) {
                    result.merges.push({
                        s: { r: currentRowIndex, c: currentColIndex },
                        e: { r: currentRowIndex + rowspan - 1, c: currentColIndex + colspan - 1 }
                    });
                }

                // 将主单元格的数据和样式填充到所有跨越的单元格
                for (let r = 0; r < rowspan; r++) {
                    for (let c = 0; c < colspan; c++) {
                        const targetRow = currentRowIndex + r;
                        const targetCol = currentColIndex + c;

                        // 确保行存在
                        while (result.data.length <= targetRow) {
                            result.data.push([]);
                            result.styles.push([]);
                        }

                        // 只在主单元格位置填充值，其他位置填空
                        result.data[targetRow][targetCol] = (r === 0 && c === 0) ? value : '';
                        result.styles[targetRow][targetCol] = style;
                    }
                }

                currentColIndex += colspan;
            });

            currentRowIndex++;
        });

        // 清理临时容器
        document.body.removeChild(tempContainer);

        parseCache.set(htmlString, result);
        return result;
    } catch (error) {
        console.error('解析HTML样式失败:', error);
        // 确保在出错时也清理临时容器
        const tempContainers = document.querySelectorAll('div[style*="position:absolute"][style*="left:-9999px"]');
        tempContainers.forEach(container => {
            if (container.parentNode) {
                document.body.removeChild(container);
            }
        });
        return null;
    }
}

/**
 * 从TD元素解析样式为Cell格式（使用计算后的样式）
 */
function parseCellStyle(td) {
    try {
        // 使用 getComputedStyle 获取应用了 CSS 类后的最终样式
        const computedStyle = window.getComputedStyle(td);
        const style = {};

        // 解析字体
        const font = {};

        // 字体粗细
        if (computedStyle.fontWeight === 'bold' || parseInt(computedStyle.fontWeight) >= 700) {
            font.bold = true;
        }

        // 字体样式
        if (computedStyle.fontStyle === 'italic') {
            font.italic = true;
        }

        // 文本装饰
        const textDecoration = computedStyle.textDecoration || '';
        if (textDecoration.includes('underline')) {
            font.underline = 'single';
        }
        if (textDecoration.includes('line-through')) {
            font.strike = true;
        }

        // 字体大小（从 pt 转换）
        if (computedStyle.fontSize) {
            const sizeInPx = parseFloat(computedStyle.fontSize);
            if (sizeInPx) {
                // 将px转换为pt（Excel标准单位）：pt = px * 0.75
                font.size = Math.round(sizeInPx * 0.75);
            }
        }

        // 字体颜色
        const color = computedStyle.color;
        if (color && color !== 'rgb(0, 0, 0)' && color !== 'rgba(0, 0, 0, 0)') {
            font.color = rgbToHex(color);
        }

        // 字体名称
        if (computedStyle.fontFamily) {
            font.name = computedStyle.fontFamily.replace(/["']/g, '').split(',')[0].trim();
        }

        if (Object.keys(font).length > 0) {
            style.font = font;
        }

        // 解析背景
        const bgColor = computedStyle.backgroundColor;
        if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
            style.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: rgbToHex(bgColor)
            };
        }

        // 解析边框
        const border = {};
        ['top', 'right', 'bottom', 'left'].forEach(side => {
            const borderStyle = computedStyle[`border${side.charAt(0).toUpperCase() + side.slice(1)}Style`];
            if (borderStyle && borderStyle !== 'none') {
                const borderColor = computedStyle[`border${side.charAt(0).toUpperCase() + side.slice(1)}Color`];
                const borderWidth = computedStyle[`border${side.charAt(0).toUpperCase() + side.slice(1)}Width`];

                border[side] = {
                    style: parseFloat(borderWidth) > 1.5 ? 'medium' : 'thin',
                    color: rgbToHex(borderColor)
                };
            }
        });

        if (Object.keys(border).length > 0) {
            style.border = border;
        }

        // 解析对齐
        const alignment = {};
        const textAlign = computedStyle.textAlign;
        if (textAlign && textAlign !== 'start' && textAlign !== 'left') {
            alignment.horizontal = textAlign;
        }

        const verticalAlign = computedStyle.verticalAlign;
        if (verticalAlign && verticalAlign !== 'baseline') {
            alignment.vertical = verticalAlign === 'middle' ? 'center' : verticalAlign;
        }

        // 换行
        const whiteSpace = computedStyle.whiteSpace;
        if (whiteSpace === 'normal' || whiteSpace === 'pre-wrap') {
            alignment.wrapText = true;
        }

        if (Object.keys(alignment).length > 0) {
            style.alignment = alignment;
        }

        return style;
    } catch (error) {
        console.error('解析单元格样式失败:', error);
        return {};
    }
}

/**
 * RGB转十六进制（优化版，使用位运算）
 */
function rgbToHex(rgb) {
    if (!rgb || rgb === 'transparent') return null;

    const match = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
    if (!match) return rgb.startsWith('#') ? rgb : null;

    const r = parseInt(match[1]);
    const g = parseInt(match[2]);
    const b = parseInt(match[3]);

    // 使用位运算优化
    return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

/**
 * HTML转义（防止XSS）
 */
function escapeHtml(text) {
    if (typeof text !== 'string') return text;

    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

/**
 * 清除缓存（在数据量大时调用，避免内存泄漏）
 */
export function clearClipboardCache() {
    styleCache.clear();
    colorCache.clear();
    parseCache.clear();
}
