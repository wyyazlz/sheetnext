// ==================== 富文本编辑器工具函数 ====================

/**
 * HTML 转义
 */
function escapeHTML(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * 根据 font 对象生成 CSS style 字符串
 * @param {Object} font - run 的 font 属性
 * @param {Object} cellFont - 单元格级别 font（fallback）
 * @param {number} zoom - 缩放比例
 * @returns {string} CSS style 字符串
 */
function fontToCSS(font, cellFont, zoom) {
    const f = { ...cellFont, ...font };
    const parts = [];
    if (f.bold) parts.push('font-weight:bold');
    else if (f.bold === false) parts.push('font-weight:normal');
    if (f.italic) parts.push('font-style:italic');
    else if (f.italic === false) parts.push('font-style:normal');
    const deco = [];
    if (f.underline) deco.push('underline');
    if (f.strike) deco.push('line-through');
    if (deco.length) parts.push(`text-decoration:${deco.join(' ')}`);
    else if (f.underline === '' || f.strike === false) parts.push('text-decoration:none');
    if (f.size) parts.push(`font-size:${f.size * 1.333 * zoom}px`);
    if (f.name) parts.push(`font-family:'${f.name}'`);
    if (f.color) parts.push(`color:${f.color}`);
    return parts.join(';');
}

/**
 * 将 _richText 数组转为 contenteditable 可显示的 HTML
 * @param {Array} runs - [{text, font}]
 * @param {Object} cellFont - 单元格级别 font
 * @param {number} zoom - 缩放比例
 * @returns {string} HTML 字符串
 */
export function richTextToHTML(runs, cellFont, zoom = 1) {
    if (!runs || !runs.length) return '';
    return runs.map(run => {
        const style = fontToCSS(run.font || {}, cellFont, zoom);
        const text = escapeHTML(run.text).replace(/\n/g, '<br>');
        return `<span style="${style}">${text}</span>`;
    }).join('');
}

/**
 * 从 DOM 节点提取 font 属性（向上查找 span/b/i/u/s 标签）
 * @param {Node} node - DOM 节点
 * @param {HTMLElement} root - 编辑器根元素
 * @returns {Object} font 对象
 */
function _getFontFromNode(node, root) {
    const font = {};
    let el = node.nodeType === 3 ? node.parentElement : node;
    while (el && el !== root) {
        // bold（内层优先，已确定则不覆盖）
        if (!('bold' in font)) {
            if (el.tagName === 'B' || el.tagName === 'STRONG') font.bold = true;
            else if (el.style?.fontWeight === 'bold' || el.style?.fontWeight >= 700) font.bold = true;
            else if (el.style?.fontWeight === 'normal') font.bold = false;
        }
        // italic
        if (!('italic' in font)) {
            if (el.tagName === 'I' || el.tagName === 'EM') font.italic = true;
            else if (el.style?.fontStyle === 'italic') font.italic = true;
            else if (el.style?.fontStyle === 'normal') font.italic = false;
        }
        // underline & strike（textDecoration 是复合属性，一次性确定）
        if ((!('underline' in font) || !('strike' in font)) && el.style) {
            const td = el.style.textDecoration || el.style.textDecorationLine || '';
            if (td) {
                if (!('underline' in font)) font.underline = td.includes('underline') ? 'single' : '';
                if (!('strike' in font)) font.strike = td.includes('line-through');
            }
        }
        if (!('underline' in font) && el.tagName === 'U') font.underline = 'single';
        if (!('strike' in font) && (el.tagName === 'S' || el.tagName === 'STRIKE' || el.tagName === 'DEL')) font.strike = true;
        // size / name / color（内层优先）
        if (el.style) {
            if (!('size' in font) && el.style.fontSize) {
                const px = parseFloat(el.style.fontSize);
                if (px > 0) font.size = Math.round(px / 1.333 * 10) / 10;
            }
            if (!('name' in font) && el.style.fontFamily) {
                font.name = el.style.fontFamily.replace(/['"]/g, '');
            }
            if (!('color' in font) && el.style.color) font.color = _normalizeColor(el.style.color);
        }
        el = el.parentElement;
    }
    return font;
}

/**
 * 将 CSS 颜色值标准化为 #RRGGBB
 */
function _normalizeColor(cssColor) {
    if (!cssColor) return '';
    if (cssColor.startsWith('#')) return cssColor;
    const m = cssColor.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (m) {
        const hex = (n) => parseInt(n).toString(16).padStart(2, '0');
        return `#${hex(m[1])}${hex(m[2])}${hex(m[3])}`;
    }
    return cssColor;
}

/**
 * 去除与 cellFont 相同的属性，保持 run.font 精简
 */
function _stripCellFont(font, cellFont) {
    const result = {};
    let has = false;
    if (font.bold && !cellFont.bold) { result.bold = true; has = true; }
    else if (font.bold === false && cellFont.bold) { result.bold = false; has = true; }
    if (font.italic && !cellFont.italic) { result.italic = true; has = true; }
    else if (font.italic === false && cellFont.italic) { result.italic = false; has = true; }
    if (font.strike && !cellFont.strike) { result.strike = true; has = true; }
    else if (font.strike === false && cellFont.strike) { result.strike = false; has = true; }
    if (font.underline && font.underline !== cellFont.underline) { result.underline = font.underline; has = true; }
    else if (font.underline === '' && cellFont.underline) { result.underline = ''; has = true; }
    if (font.size && font.size != (cellFont.size || 11)) { result.size = font.size; has = true; }
    if (font.name && font.name !== (cellFont.name || '宋体')) { result.name = font.name; has = true; }
    if (font.color && font.color !== cellFont.color) { result.color = font.color; has = true; }
    return has ? result : {};
}

/**
 * 递归遍历 DOM 收集 runs
 */
function _walkNodes(node, root, cellFont, runs, prevIsBlock) {
    for (let child = node.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 3) {
            // 文本节点
            const text = child.textContent;
            if (text) {
                const font = _stripCellFont(_getFontFromNode(child, root), cellFont);
                runs.push({ text, font });
            }
        } else if (child.nodeType === 1) {
            const tag = child.tagName;
            if (tag === 'BR') {
                // 跳过末尾占位 br
                if (!child.nextSibling || (child.nextSibling.nodeType === 1 && child.nextSibling.tagName === 'BR' && !child.nextSibling.nextSibling)) {
                    runs.push({ text: '\n', font: {} });
                    break;
                }
                runs.push({ text: '\n', font: {} });
            } else if (tag === 'DIV' || tag === 'P') {
                // 块级元素前插入换行（除非是第一个子元素）
                if (runs.length > 0 && runs[runs.length - 1].text !== '\n') {
                    runs.push({ text: '\n', font: {} });
                }
                _walkNodes(child, root, cellFont, runs, true);
            } else {
                _walkNodes(child, root, cellFont, runs, false);
            }
        }
    }
}

/**
 * 合并相邻同格式 run
 */
function _mergeRuns(runs) {
    if (!runs.length) return runs;
    const merged = [runs[0]];
    for (let i = 1; i < runs.length; i++) {
        const prev = merged[merged.length - 1];
        const cur = runs[i];
        if (JSON.stringify(prev.font) === JSON.stringify(cur.font)) {
            prev.text += cur.text;
        } else {
            merged.push(cur);
        }
    }
    return merged;
}

/**
 * 解析 contenteditable DOM 树，还原为 _richText 数组
 * @param {HTMLElement} inputElement - 编辑器元素
 * @param {Object} cellFont - 单元格级别 font
 * @returns {Array|null} richText 数组，或 null（纯文本）
 */
export function htmlToRichText(inputElement, cellFont) {
    const runs = [];
    _walkNodes(inputElement, inputElement, cellFont, runs, false);
    if (!runs.length) return null;
    const merged = _mergeRuns(runs);
    // 去除末尾空换行
    while (merged.length && merged[merged.length - 1].text === '\n') merged.pop();
    if (!merged.length) return null;
    // 如果只有一个 run 且 font 为空，返回 null（纯文本）
    if (merged.length === 1 && Object.keys(merged[0].font).length === 0) return null;
    return merged;
}

/**
 * 收集 range 内所有文本节点
 */
function _getTextNodesInRange(range, root) {
    const nodes = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
        if (range.intersectsNode(node)) nodes.push(node);
    }
    return nodes;
}

/**
 * 获取当前选区/光标位置的字体属性
 * @param {HTMLElement} inputElement - 编辑器元素
 * @param {Object} cellFont - 单元格级别 font
 * @returns {Object} font 对象（含 bold, italic, underline, strike, size, name, color）
 */
export function getSelectionFont(inputElement, cellFont) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return { ...cellFont };
    const range = sel.getRangeAt(0);
    if (!inputElement.contains(range.commonAncestorContainer)) return { ...cellFont };

    if (range.collapsed) {
        const raw = _getFontFromNode(range.startContainer, inputElement);
        return { ...(cellFont || {}), ...raw };
    }

    // 有选区：取所有文本节点的公共 font
    const textNodes = _getTextNodesInRange(range, inputElement);
    if (!textNodes.length) return { ...cellFont };

    const fonts = textNodes.map(n => {
        const raw = _getFontFromNode(n, inputElement);
        return { ...(cellFont || {}), ...raw };
    });

    const common = { ...fonts[0] };
    for (let i = 1; i < fonts.length; i++) {
        for (const key of Object.keys(common)) {
            if (common[key] !== fonts[i][key]) {
                delete common[key];
            }
        }
    }
    return common;
}

// toggle 类属性
const TOGGLE_PROPS = new Set(['bold', 'italic', 'strike', 'underline']);

/**
 * 为 span 设置单个 font 属性对应的 CSS
 */
function _setSpanStyle(span, prop, value) {
    if (prop === 'bold') {
        span.style.fontWeight = value ? 'bold' : 'normal';
    } else if (prop === 'italic') {
        span.style.fontStyle = value ? 'italic' : 'normal';
    } else if (prop === 'underline') {
        const strike = (span.style.textDecoration || '').includes('line-through');
        const parts = [];
        if (value) parts.push('underline');
        if (strike) parts.push('line-through');
        span.style.textDecoration = parts.length ? parts.join(' ') : 'none';
    } else if (prop === 'strike') {
        const ul = (span.style.textDecoration || '').includes('underline');
        const parts = [];
        if (ul) parts.push('underline');
        if (value) parts.push('line-through');
        span.style.textDecoration = parts.length ? parts.join(' ') : 'none';
    } else if (prop === 'size') {
        span.style.fontSize = value ? `${value * 1.333}px` : '';
    } else if (prop === 'name') {
        span.style.fontFamily = value || '';
    } else if (prop === 'color') {
        span.style.color = value || '';
    }
}

/**
 * 对选中文字应用格式
 * @param {HTMLElement} inputElement - 编辑器元素
 * @param {string} fontProp - 属性名（bold/italic/strike/underline/size/name/color）
 * @param {*} value - 属性值
 */
export function applyFormatToSelection(inputElement, fontProp, value) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);
    if (!inputElement.contains(range.commonAncestorContainer)) return;

    // toggle 属性：检测当前状态后取反
    // 注意：range.startContainer 可能是元素节点（如 selectNodeContents 设置的选区），
    // 需要找到实际的文本节点才能正确检测当前样式
    if (TOGGLE_PROPS.has(fontProp) && typeof value === 'boolean') {
        let refNode = range.startContainer;
        if (refNode.nodeType !== 3) {
            const textNodes = _getTextNodesInRange(range, inputElement);
            if (textNodes.length) refNode = textNodes[0];
        }
        const curFont = _getFontFromNode(refNode, inputElement);
        if (fontProp === 'underline') {
            value = curFont.underline ? '' : 'single';
        } else {
            value = !curFont[fontProp];
        }
    }

    if (range.collapsed) {
        // 无选区：插入零宽 span 使后续输入继承格式
        const span = document.createElement('span');
        _setSpanStyle(span, fontProp, value);
        // 继承当前位置的其他样式
        const parentSpan = range.startContainer.nodeType === 3
            ? range.startContainer.parentElement : range.startContainer;
        if (parentSpan && parentSpan !== inputElement && parentSpan.style) {
            const cs = parentSpan.style;
            if (!span.style.fontWeight && cs.fontWeight) span.style.fontWeight = cs.fontWeight;
            if (!span.style.fontStyle && cs.fontStyle) span.style.fontStyle = cs.fontStyle;
            if (!span.style.fontSize && cs.fontSize) span.style.fontSize = cs.fontSize;
            if (!span.style.fontFamily && cs.fontFamily) span.style.fontFamily = cs.fontFamily;
            if (!span.style.color && cs.color) span.style.color = cs.color;
            if (!span.style.textDecoration && cs.textDecoration) span.style.textDecoration = cs.textDecoration;
        }
        span.textContent = '\u200B'; // 零宽空格
        range.insertNode(span);
        const newRange = document.createRange();
        newRange.setStart(span.firstChild, 1);
        newRange.collapse(true);
        sel.removeAllRanges();
        sel.addRange(newRange);
        return;
    }

    // 有选区：提取内容，包裹新 span
    const fragment = range.extractContents();
    const wrapper = document.createElement('span');

    _applyFormatToFragment(fragment, inputElement, fontProp, value);
    wrapper.appendChild(fragment);

    range.insertNode(wrapper);

    // 清理：移除仅包含 wrapper 的空祖先 span（避免残留的旧样式通过继承干扰）
    // 注意：extractContents() 可能留下空文本节点，不能用 childNodes.length === 1 判断
    let parent = wrapper.parentElement;
    while (parent && parent !== inputElement && parent.tagName === 'SPAN') {
        const hasOtherContent = [...parent.childNodes].some(n =>
            n !== wrapper && (n.nodeType === 1 || (n.nodeType === 3 && n.textContent))
        );
        if (!hasOtherContent) {
            parent.parentElement.insertBefore(wrapper, parent);
            parent.remove();
            parent = wrapper.parentElement;
        } else {
            break;
        }
    }

    // 恢复选区
    const newRange = document.createRange();
    newRange.selectNodeContents(wrapper);
    sel.removeAllRanges();
    sel.addRange(newRange);
}

/**
 * 对 fragment 中的所有 span 应用格式
 */
function _applyFormatToFragment(fragment, root, prop, value) {
    const spans = fragment.querySelectorAll('span[style]');
    if (spans.length) {
        spans.forEach(span => _setSpanStyle(span, prop, value));
    }
    // 处理直接文本节点（无 span 包裹的）
    for (let child = fragment.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === 3 && child.textContent) {
            const span = document.createElement('span');
            _setSpanStyle(span, prop, value);
            span.textContent = child.textContent;
            fragment.replaceChild(span, child);
            child = span;
        }
    }
}

/**
 * 快速检测编辑器内是否含格式化内容
 * @param {HTMLElement} inputElement
 * @returns {boolean}
 */
export function hasRichContent(inputElement) {
    return inputElement.querySelector('span[style], b, i, u, s, strong, em, strike, del') !== null;
}
